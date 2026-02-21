import adsk.core, adsk.fusion, adsk.cam, traceback
import re
import os, itertools

_app = None
_ui = None
_handlers = []

CMD_ID   = 'batchParametricExportCmd'
CMD_NAME = 'Batch parametric export'
CMD_DESC = 'Select bodies and components to export and process them using all parameter combinations.'

_param_registry = {}  # key: chk_id -> {'param': adsk.fusion.UserParameter, 'text_id': str}

_SIMPLE_LITERAL_RE = re.compile(r'^\s*([-+]?\d+(?:\.\d+)?)\s*([A-Za-z%°/]+)?\s*$')
_TEXT_LITERAL_RE = re.compile( r'^\s*\'([^\']*)\'\s*$' )# re.compile(r'^\s*\'(.+?)\'\s*$')

# --- filename template wiring ---
FORMAT_ID = 'outFormat'
FILENAME_ID = 'filenameTemplate'
_current_inputs = None  # set during command; used by change handler

# keep parameter order for template construction
_param_registry = {}   # chk_id -> {'param': UserParameter, 'text_id': str, 'name': str, 'order': int}
_param_order = []      # list of parameter names in UI order

_EXT_MAP = {'STEP': 'step', 'STL': 'stl', '3MF': '3mf', 'OBJ': 'obj'}

OUTPUT_DIR_ID = 'outputDir'
OUTPUT_BROWSE_ID = 'outputBrowse'
_last_folder = ''  # remember the last chosen folder during the session

# name -> ('body'|'component', ref)
_item_registry = {}

_orig_body_vis = None
_orig_occ_vis = None
_prog = None
_step = 0

def run(context):
    try:
        global _app, _ui
        _app = adsk.core.Application.get()
        _ui  = _app.userInterface

        cmd_def = _ui.commandDefinitions.itemById(CMD_ID)
        if not cmd_def:
            cmd_def = _ui.commandDefinitions.addButtonDefinition(CMD_ID, CMD_NAME, CMD_DESC)

        class CommandCreatedHandler(adsk.core.CommandCreatedEventHandler):
            def notify(self, args: adsk.core.CommandCreatedEventArgs):
                try:
                    cmd = args.command
                    cmd.setDialogMinimumSize(1000, 400)
                    cmd.setDialogInitialSize(1000, 400)
                    cmd.okButtonText = 'Process'  # OK = Process
                    inputs = cmd.commandInputs
                    
                    # We don't need this anymore, will be handled through selection sets instead
                    #_createObjectsTable(inputs)

                    _createParametersTable(inputs)

                    # --- NEW: output format dropdown ---
                    #fmt = inputs.addDropDownCommandInput(FORMAT_ID, 'Output format',
                    #                                    adsk.core.DropDownStyles.TextListDropDownStyle)
                    #for opt in ('STEP', 'STL', '3MF', 'OBJ'):
                    #    fmt.listItems.add(opt, opt == 'STL', '')  # default STL selected

                    # --- NEW: filename template input ---
                    tmpl = inputs.addStringValueInput(FILENAME_ID, 'Filename template', '')
                    tmpl.tooltip = 'Uses {name} and selected parameter names. Example: {name}_{width}_{height}.obj'
                    


                    # wire change handler so template updates
                    change_handler = InputChangedHandler()
                    cmd.inputChanged.add(change_handler)
                    _handlers.append(change_handler)

                    # stash inputs for the change handler
                    global _current_inputs
                    _current_inputs = inputs

                    _createOutputFolderPicker(inputs)

                    # Events
                    exec_handler = ExecuteHandler()
                    destroy_handler = CommandDestroyedHandler()
                    cmd.execute.add(exec_handler)
                    cmd.destroy.add(destroy_handler)
                    _handlers.extend([exec_handler, destroy_handler])

                    validate_handler = ValidateHandler()
                    cmd.validateInputs.add(validate_handler)
                    _handlers.append(validate_handler)

                    design = adsk.fusion.Design.cast(_app.activeProduct)
                    if not design:
                        _app.log('[BatchExport] No active design.')
                        return

                    ups = design.userParameters
                    if not ups or ups.count == 0:
                        _app.log('[BatchExport] No user parameters found.')
                        return

                    _app.log('[BatchExport] Non-formula user parameters:')
                    for i in range(ups.count):
                        p = ups.item(i)
                        if p.name == "output_filename":
                            tmpl.value = p.textValue
                            continue

                        expr = p.expression  # what you see in Parameters dialog
                        if _is_simple_literal(expr):
                            # Log: Name, expression, evaluated value in internal units, and unit
                            _app.log(f'  - {p.name}: expression="{expr}", value={p.value}, unit="{p.unit}"')

                            # Show dialog
                            cmd.doExecute(False)

                except:
                    _ui.messageBox('Command create failed:\n{}'.format(traceback.format_exc()))

        class ValidateHandler(adsk.core.ValidateInputsEventHandler):
            def notify(self, args: adsk.core.ValidateInputsEventArgs):
                try:
                    
                   
                    #_ui.messageBox(f'test: {}')
                    design = adsk.core.Application.get().activeProduct
                    if not design:
                        _ui.messageBox('[BatchExport] No active design.')
                        return

                    ok, msg = _validate_all(args.inputs, design)


                    args.areInputsValid = ok
                    # Optional: show inline error text (no popups)
                    err = adsk.core.TextBoxCommandInput.cast(args.inputs.itemById('inline_error'))
                    if not err:
                        # create it once (read-only, full width)
                        err = args.inputs.addTextBoxCommandInput('inline_error', '', '', 2, True)
                        err.isFullWidth = True
                    err.text = '' if ok else f'{msg}'
                    err.isVisible = not ok
                except:
                    # if validation itself fails, keep OK disabled
                    args.areInputsValid = False

        class ExecuteHandler(adsk.core.CommandEventHandler):
            def notify(self, args: adsk.core.CommandEventArgs):
                try:
                    global _prog, _step, _orig_body_vis, _orig_occ_vis
                    design = adsk.fusion.Design.cast(_app.activeProduct)
                    if not design:
                        _ui.messageBox('No active design.')
                        return


                    parameters = design.userParameters
                    inputs = args.command.commandInputs
                    template = adsk.core.StringValueCommandInput.cast(inputs.itemById(FILENAME_ID)).value.strip()
                    out_dir = adsk.core.StringValueCommandInput.cast(inputs.itemById(OUTPUT_DIR_ID)).value.strip()
                    
                    selectionSets = design.selectionSets
                    if selectionSets.count == 0:
                        _ui.messageBox('No selection sets defined.')
                        return
                    
                    numStaticModels = 0
                    numParametricModels = 0
                    for selectionSet in selectionSets:
                        # All bodies in the static sets are exported without applying any parameters
                        if selectionSet.name.startswith('Static_'):
                            for entity in selectionSet.entities:
                                if isinstance(entity, adsk.fusion.BRepBody):
                                    numStaticModels += 1
                        # All bodies in the parametric sets are exported with applied parameters
                        if selectionSet.name.startswith('Parametric_'):
                            for entity in selectionSet.entities:
                                if isinstance(entity, adsk.fusion.BRepBody):
                                    numParametricModels += 1

                    formats = None
                    for p in parameters:
                        if p.name == 'output_formats':
                            formats = p.textValue.split(';')
                            break
                    if not formats:
                        _ui.messageBox("Missing parameter 'output_formats'")
                        return


                    sel_param_names, param_values = _get_selected_params_and_values(inputs, design)
                     # materialize all combos once (preserves UI order)
                    ordered_lists = [param_values[p] for p in sel_param_names]
                    all_combos = list(itertools.product(*ordered_lists))

                    total = (len(all_combos) * numParametricModels + numStaticModels) * len(formats) + 1

                    _prog = _progress_start('Batch export', total)
                    _step = 0

                    orig_param_expr = {}
                    _orig_body_vis, _orig_occ_vis = _snapshot_visibility(design)

                    try:
                        for selectionSet in selectionSets:
                            # All bodies in the static sets are exported without applying any parameters
                            if selectionSet.name.startswith('Static_'):
                                selectionSetName = selectionSet.name[7:]
                                _app.log(f"[BatchExport] Found static view set '{selectionSetName}'")

                                _exortBudies(selectionSet, template, formats, out_dir, selectionSetName, design)

                            # All bodies in the parametric sets are exported with applied parameters
                            if selectionSet.name.startswith('Parametric_'):
                                selectionSetName = selectionSet.name[11:]
                                _app.log(f"[BatchExport] Found parametric view set '{selectionSetName}'")

                                for combo in all_combos:
                                    _set_user_params(design, param_values, combo, sel_param_names, orig_param_expr)
                                    _compute(design)

                                    # reusable mapping for filenames
                                    pv_map = dict(zip(sel_param_names, combo))

                                    _exortBudies(selectionSet, template, formats, out_dir, selectionSetName, design, pv_map)

                        document = adsk.fusion.FusionDocument.cast(_app.activeDocument)
                        out_filename = _sanitize_filename_component(document.name)+'.f3d'
                        _progress_update(f'{out_filename}')
                        exportManager = design.exportManager
                        exportOptions = exportManager.createFusionArchiveExportOptions(os.path.join(out_dir, out_filename))
                        exportManager.execute(exportOptions)

                    except KeyboardInterrupt:
                        _ui.messageBox('Export cancelled.')
                    finally:
                        _progress_end(_prog)
                        _restore_user_params(design, orig_param_expr)
                        _restore_visibility(design)
                        _compute(design)

                        adsk.terminate()

                except Exception as ex:
                    _ui.messageBox('Execute failed:\n{}'.format(traceback.format_exc()))

        class CommandDestroyedHandler(adsk.core.CommandEventHandler):
            def notify(self, args):
                _item_registry.clear()
                _param_registry.clear()
                global _current_inputs
                _current_inputs = None

        class InputChangedHandler(adsk.core.InputChangedEventHandler):
            def notify(self, args: adsk.core.InputChangedEventArgs):
                global _last_folder
                try:
                    inp = args.input
                    if not inp:
                        return

                    # Browse for folder
                    if inp.id == OUTPUT_BROWSE_ID and inp.value:  # clicked
                        dlg = _ui.createFolderDialog()
                        dlg.title = 'Select output folder'
                        if _last_folder:
                            dlg.initialDirectory = _last_folder
                        res = dlg.showDialog()
                        if res == adsk.core.DialogResults.DialogOK:
                            folder = dlg.folder
                            sv = adsk.core.StringValueCommandInput.cast(_current_inputs.itemById(OUTPUT_DIR_ID))
                            if sv:
                                sv.value = folder
                            _last_folder = folder

                    
                            
                        # reset so button can be clicked again
                        inp.value = False

                except:
                    _ui.messageBox('inputChanged failed:\n{}'.format(traceback.format_exc()))


        created_handler = CommandCreatedHandler()
        cmd_def.commandCreated.add(created_handler)
        _handlers.append(created_handler)

        cmd_def.execute()
        adsk.autoTerminate(False)

    except:
        if _ui:
            _ui.messageBox('Add-in run failed:\n{}'.format(traceback.format_exc()))
            adsk.terminate()
            
def _exortBudies(selectionSet : adsk.core.SelectionSet, template : str, formats : list[str], out_dir : str, selectionSetName: str, design : adsk.fusion.Design, pv_map: dict = dict()):
    global _step
    for entity in selectionSet.entities:
        # only consider bodies for the export
        if isinstance(entity, adsk.fusion.BRepBody):
            body = adsk.fusion.BRepBody.cast(entity)

            out_filename = _build_filename(template, body.name, selectionSetName, pv_map)
            
            for format in formats:
                #seletionSetFolder = os.path.join(out_dir, selectionSetName)
                fullpath = os.path.join(out_dir, out_filename + "." + format)

                if not os.path.exists(os.path.dirname(fullpath)):
                    os.makedirs(os.path.dirname(fullpath), exist_ok=True)
                else:
                    if not os.path.isdir(os.path.dirname(fullpath)):
                        _ui.messageBox(f'Unable to create selection set folder: {os.path.dirname(fullpath)}')
                        return
                
                if pv_map:
                    combo_note = '\n'.join(f'{p}={pv_map[p]}' for p in pv_map)
                    _progress_update(f'{body.name} \n{combo_note} \n\n{out_filename}.{format}')
                else:
                    _progress_update(f'{body.name} \n\n{out_filename}.{format}')

                if format.upper() == 'STEP':
                    # isolate just this object for STEP
                    _isolate_for_step(design, 'body', body)
                    # No need to recompute geometry; visibility changes don’t require it
                    ok = _export_step(design, fullpath)
                    # restore vis right after exporting this object to keep scene sane
                    _restore_visibility(design)
                else:
                    # mesh formats target the entity directly
                    ok = _export_mesh(design, body, fullpath, format)

                if not ok:
                    raise RuntimeError(f'Export failed: {fullpath}')
                
                _step +=1

def stop(context):
    try:
        global _ui
        if _ui:
            cmd_def = _ui.commandDefinitions.itemById(CMD_ID)
            if cmd_def:
                cmd_def.deleteMe()
    except:
        if _ui:
            _ui.messageBox('Add-in stop failed:\n{}'.format(traceback.format_exc()))

def _createObjectsTable(inputs: adsk.core.CommandInputs):
    # Caption (read-only)
    caption = inputs.addTextBoxCommandInput(
        'caption', 'Batch parametric export',
        'Select bodies and components to export', 1, True
    )
    caption.isFullWidth = True

    table = inputs.addTableCommandInput('itemsTable', 'Select bodies and components to export', 2, '1:8')
    table.columnSpacing = 6
    table.rowSpacing = 2
    # table.hasGrid = False  # optional: cleaner look

    design = adsk.fusion.Design.cast(_app.activeProduct)
    if not design:
        return
    root = design.rootComponent
    occs = root.occurrences

    row = 0

    # Root bodies
    for i in range(root.bRepBodies.count):
        b = root.bRepBodies.item(i)
        chk_id = f'chk_body_{i}'
        chk = inputs.addBoolValueInput(chk_id, '', True, '', b.isVisible)
        lbl = inputs.addTextBoxCommandInput(f'lbl_body_{i}', '', f'[Body] {b.name}', 1, True)
        table.addCommandInput(chk, row, 0)
        table.addCommandInput(lbl, row, 1)
        _item_registry[chk_id] = ('body', b)
        row += 1

    # Components (skip root)
    for j in range(occs.count):
        occ = occs.item(j)
        if occ.component == root:
            continue
        comp = occ.component
        chk_id = f'chk_comp_{j}'
        chk = inputs.addBoolValueInput(chk_id, '', True, '', True)
        lbl = inputs.addTextBoxCommandInput(f'lbl_comp_{j}', '', f'[Comp] {comp.name}', 1, True)
        table.addCommandInput(chk, row, 0)
        table.addCommandInput(lbl, row, 1)
        _item_registry[chk_id] = ('component', occ)
        row += 1

def _is_simple_literal(expr: str) -> bool:
    return bool(expr and _SIMPLE_LITERAL_RE.match(expr))

def _is_text_literal(expr: str) -> bool:
    return bool(expr and _TEXT_LITERAL_RE.match(expr))

def _format_expr_2dec(expr: str) -> str:
    """
    Format a parameter expression like '12 mm' into '12.00 mm'.
    If no unit, returns '12.00'. If it isn't a simple literal, just return expr.
    """
    m = _SIMPLE_LITERAL_RE.match(expr or '')
    if not m:
        return expr or ''
    num = float(m.group(1))
    unit = (m.group(2) or '').strip()
    return f'{num:.2f}{" " + unit if unit else ""}'

def _safe_id(text: str) -> str:
    return ''.join(c if c.isalnum() else '_' for c in (text or ''))

def _createParametersTable(inputs: adsk.core.CommandInputs):
    """3 columns: [Name][Current Value][Values to iterate]; non-formula user params only."""
    design = adsk.fusion.Design.cast(_app.activeProduct)
    if not design:
        return

    # instruction
    #hint = inputs.addTextBoxCommandInput(
    #    'params_hint', 'Select parameters to iterate',
    #    'Enter values as semicolon-separated list, e.g.: 1; 5.5; 12; 15', 2, True
    #)
    #hint.isFullWidth = True

    hint = inputs.addTextBoxCommandInput('params_heading', '', 'List of detected parameters and variants:', 1, True)
    hint.isFullWidth = True

    tbl = inputs.addTableCommandInput('paramsTable', 'Select parameters to iterate', 4, '4:3:8')
    tbl.columnSpacing = 6
    tbl.rowSpacing = 2
    # Optional scroller:
    tbl.minimumVisibleRows = 1
    # tbl.maximumVisibleRows = 12

    tbl.addCommandInput(inputs.addTextBoxCommandInput("head_parameter", '', "Parameter", 1, True), 0, 0)
    tbl.addCommandInput(inputs.addTextBoxCommandInput("head_value", '', "Value", 1, True), 0, 1)
    tbl.addCommandInput(inputs.addTextBoxCommandInput("head_variants", '', "Variants", 1, True), 0, 2)

    ups = design.userParameters
    if not ups or ups.count == 0:
        info = inputs.addTextBoxCommandInput('no_params', '', 'No user parameters found.', 1, True)
        tbl.addCommandInput(info, 0, 0)
        return

    # clear prior order if re-entered within same session
    _param_order.clear()

    row = 1
    for i in range(ups.count):
        p = ups.item(i)
        if not _is_simple_literal(p.expression) and not _is_text_literal(p.expression):
            continue
        if not p.isFavorite:
            continue
        if not p.comment:
            continue

        base = _safe_id(p.name)
        chk_id  = f'chk_param_{base}'
        name_id = f'lbl_param_{base}'
        val_id  = f'lbl_value_{base}'
        txt_id  = f'txt_values_{base}'

        #chk = inputs.addBoolValueInput(chk_id, '', True, '', False)
        #tbl.addCommandInput(chk, row, 0)

        name_lbl = inputs.addTextBoxCommandInput(name_id, '', p.name, 1, True)
        tbl.addCommandInput(name_lbl, row, 0)

        cur_val = _format_expr_2dec(p.expression) or p.expression or ''
        val_lbl = inputs.addTextBoxCommandInput(val_id, '', cur_val, 1, True)
        tbl.addCommandInput(val_lbl, row, 1)


        txt = inputs.addStringValueInput(txt_id, '', p.comment)
        txt.tooltip = 'Semicolon-separated values, e.g.: 1; 5.5; 12; 15'
        tbl.addCommandInput(txt, row, 2)

        _param_registry[chk_id] = {'param': p, 'text_id': txt_id, 'name': p.name, 'order': row}
        _param_order.append(p.name)
        row += 1

    if row == 0:
        info = inputs.addTextBoxCommandInput('no_simple_params', '', 'No non-formula parameters found.', 1, True)
        tbl.addCommandInput(info, 0, 0)

def _createOutputFolderPicker(inputs: adsk.core.CommandInputs):
    # Two-column table: [path field][Browse…]
    #tab = inputs.addTabCommandInput('outDirTable', 'Output folder')

    #tbl.columnSpacing = 6
    #tbl.rowSpacing = 2

    #path_in = inputs.addStringValueInput(OUTPUT_DIR_ID, 'Output folder', _last_folder or '')
    #path_in.isFullWidth = True
    #tab.children.addCommandInput(path_in, 0, 0)
    #tab.children.addStringValueInput(OUTPUT_DIR_ID, 'Output folder', _last_folder or '')
    inputs.addStringValueInput(OUTPUT_DIR_ID, 'Output folder', _last_folder or '')

    # br = inputs.addButtonRowCommandInput(OUTPUT_BROWSE_ID, '', False)  # no multi-select
    # br.listItems.add('Browse…', False, '')
    # tbl.addCommandInput(br, 0, 1)

    #browseBtn = inputs.addBoolValueInput(OUTPUT_BROWSE_ID, 'Browse…', False, '', False)
    #tab.children.addBoolValueInput(OUTPUT_BROWSE_ID, 'Browse…', False, '', False)
    inputs.addBoolValueInput(OUTPUT_BROWSE_ID, 'Browse…', False, '', False)
    #tbl.addCommandInput(browseBtn, 0, 1)

def _get_selected_objects(inputs: adsk.core.CommandInputs, design: adsk.fusion.Design):
    """Return list of tuples: [('body'|'component', obj, display_name)] based on table checkboxes."""
    selected = []
    for chk_id, (kind, ref) in _item_registry.items():
        chk = adsk.core.BoolValueCommandInput.cast(inputs.itemById(chk_id))
        if chk and chk.value:
            if kind == 'body':
                name = getattr(ref, 'name', 'Body')
            else:  # component (occurrence)
                comp = getattr(ref, 'component', None)
                name = getattr(comp, 'name', 'Component')
            selected.append((kind, ref, name))

    
    return selected

def _parse_values_list(raw: str):
    """Parse '1; 5.5; 12' into [1.0, 5.5, 12.0]. Empty strings are ignored. Raises ValueError on bad token."""
    vals = []
    if not raw:
        raise ValueError('empty')
    for tok in (t.strip() for t in raw.split(';')):
        if not tok:
            continue
        vals.append(float(tok))
    if not vals:
        raise ValueError('no numbers')
    return vals

def _parse_text_values_list(raw: str):
    # split raw by semicolon, but only when we're outside of quotes
    vals = []
    if not raw:
        raise ValueError('empty')
    
    # Split by semicolons that are outside of single quotes
    parts = []
    current_part = ""
    in_quotes = False
    i = 0
    
    while i < len(raw):
        char = raw[i]
        if char == "'" and (i == 0 or raw[i-1] != "\\"):
            in_quotes = not in_quotes
            current_part += char
        elif char == ";" and not in_quotes:
            parts.append(current_part)
            current_part = ""
        else:
            current_part += char
        i += 1
    
    if current_part:
        parts.append(current_part)
    
    for tok in (t.strip() for t in parts):
        if not tok:
            continue
        m = _TEXT_LITERAL_RE.match(tok)
        if not m:
            raise ValueError(f'bad text token: {tok}')
        vals.append(m.group(1))
    
    if not vals:
        raise ValueError('no text values')
    return vals

def _get_selected_params_and_values(inputs: adsk.core.CommandInputs, design: adsk.fusion.Design):
    """
    Returns:
      ordered_names: [ 'width', 'height', ... ] in UI order for checked items
      values_map: { 'width': [..floats..], ... }
    Raises ValueError on any unparseable list.
    """
    ups = design.userParameters
    ordered_names = []
    values_map = {}
    # _param_order keeps UI row order
    for pname in _param_order:
        # find checkbox record for this parameter
        chk_id = None
        text_id = None
        for cid, meta in _param_registry.items():
            if meta['name'] == pname:
                chk_id = cid
                text_id = meta['text_id']
                break
        if not chk_id:
            continue


        txt = adsk.core.StringValueCommandInput.cast(inputs.itemById(text_id))
        #_app.log(f"[BatchExport] Processing parameter '{pname}' with input values: '{(txt.value or '').strip()}'")
        upsItem = ups.itemByName(pname)
        if not upsItem:
            _app.log(f"[BatchExport] Warning: Parameter '{pname}' not found in design.")
        if ( upsItem.valueType == 1 ):   # text
            #_app.log(f"[BatchExport] Parameter '{pname}' is of type text.")
            vals = _parse_text_values_list((txt.value or '').strip())
            _app.log("[BatchExport] Parsed text values for parameter '{}': {}".format(pname, vals))
        else:
            #_app.log(f"[BatchExport] Parameter '{pname}' is of type number.")
            vals = _parse_values_list((txt.value or '').strip())
        
        ordered_names.append(pname)
        values_map[pname] = vals
    
    return ordered_names, values_map

def _validate_filename_template(template: str, selected_param_names):
    """Ensure {name} present and placeholders for every selected parameter. Return normalized template."""
    # TBD: more checking required for new placeholders
    if not template:
        return False, 'Filename template is empty.'
    if '{body_name}' not in template:
        return False, 'Filename template must include {body_name}.'
    #missing = [p for p in selected_param_names if f'{{{p}}}' not in template]
    #if missing:
    #    return False, 'Filename template is missing placeholders: ' + ', '.join(f'{{{p}}}' for p in missing)
    return True, ''

def _sanitize_filename_component(s: str):
    """Remove characters illegal on common filesystems."""
    return ''.join(c for c in s if c not in '\\/:*?"<>|\n\r\t').strip().replace(' ', '_')

def _build_filename(template: str, obj_name: str, selection_set_name: str, param_values_map: dict = dict()):
    document = adsk.fusion.FusionDocument.cast(_app.activeDocument)

    out = template
    out = out.replace('{project_name}', _sanitize_filename_component(document.name))
    out = out.replace('{selection_set}', _sanitize_filename_component(selection_set_name))
    out = out.replace('{body_name}', _sanitize_filename_component(obj_name))

    if len(param_values_map) > 0:
        param_strings = []
        for pname, val in param_values_map.items():
            param_strings.append(str(val));
        
        out = out.replace('{param_values}', _sanitize_filename_component('_'.join(param_strings)))
    else:
        out = out.replace('_{param_values}', _sanitize_filename_component(''))
        out = out.replace('{param_values}', _sanitize_filename_component(''))

    return out

def _normalize_path(p: str) -> str:
    # strip quotes/spaces, expand ~, resolve . and ..
    return os.path.normpath(os.path.expanduser(p.strip().strip('"').strip("'")))

def _validate_all(inputs: adsk.core.CommandInputs, design: adsk.fusion.Design):
    # params
    try:
        sel_param_names, param_values = _get_selected_params_and_values(inputs, design)
    except ValueError:
        return False, 'All selected parameters must have semicolon-separated constant values.'
        
    
    #if not sel_param_names:
    #    return False, 'Select at least one parameter.'

    # format + template
    #fmt = adsk.core.DropDownCommandInput.cast(inputs.itemById(FORMAT_ID))
    #fmt_name = fmt.selectedItem.name if (fmt and fmt.selectedItem) else 'OBJ'
    #ext = _EXT_MAP.get(fmt_name, 'obj')
    

    tmpl_in = adsk.core.StringValueCommandInput.cast(inputs.itemById(FILENAME_ID))
    template = (tmpl_in.value or '').strip() if tmpl_in else ''
    ok, msg = _validate_filename_template(template, sel_param_names)
    if not ok:
        return False, msg



    # output dir
    out_dir_in = adsk.core.StringValueCommandInput.cast(inputs.itemById(OUTPUT_DIR_ID))
    raw = (out_dir_in.value or '') if out_dir_in else ''

    if not raw:
        return False, 'Output folder cannot be empty.'
    else:
        out_dir = _normalize_path(raw)

        if not out_dir or not os.path.isdir(out_dir):
            return False, 'Output folder must exist.'
        if not os.access(out_dir, os.W_OK):
            return False, 'Output folder is not writable.'

    return True, ''

def _set_user_params(design, name_to_values_dict, values_tuple, ordered_names, originals_cache):
    """Set selected user parameters for one combo. Cache original expressions for later restore."""
    ups = design.userParameters
    for pname, pval in zip(ordered_names, values_tuple):
        up = ups.itemByName(pname)
        if not up:
            continue
        if pname not in originals_cache:
            originals_cache[pname] = up.expression  # keep original string
        unit = (up.unit or '').strip()
        if (up.valueType == 1):  # text
            _app.log(f"[BatchExport] Setting text parameter '{pname}' to '{pval.strip()}'")
            up.textValue = f"'{pval.strip()}'"
        else:
            up.expression = f'{pval} {unit}'.strip()
    return True

def _restore_user_params(design, originals_cache):
    ups = design.userParameters
    for pname, expr in originals_cache.items():
        up = ups.itemByName(pname)
        if up:
            up.expression = expr

def _compute(design):
    try:
        design.computeAll()  # Force recompute (same as Compute All)
    except:
        pass  # don’t hard-fail; export might still succeed

def _snapshot_visibility(design):
    """Return dicts of initial visibility for bodies and occurrences."""
    root = design.rootComponent
    body_vis = {}
    for i in range(root.bRepBodies.count):
        b = root.bRepBodies.item(i)
        body_vis[b.entityToken] = b.isVisible
    occ_vis = {}
    occs = root.occurrences
    for i in range(occs.count):
        oc = occs.item(i)
        occ_vis[oc.entityToken] = oc.isLightBulbOn
    return body_vis, occ_vis

def _restore_visibility(design):
    root = design.rootComponent
    for i in range(root.bRepBodies.count):
        b = root.bRepBodies.item(i)
        vis = _orig_body_vis.get(b.entityToken, True)
        try: b.isVisible = vis
        except: pass
    occs = root.occurrences
    for i in range(occs.count):
        oc = occs.item(i)
        vis = _orig_occ_vis.get(oc.entityToken, True)
        try: oc.isLightBulbOn = vis
        except: pass

def _isolate_for_step(design, kind, ref):
    """
    Hide everything except the target.
    kind: 'body' or 'component' (occurrence)
    ref:  BRepBody or Occurrence
    """
    root = design.rootComponent
    # Hide all occurrences first.
    for i in range(root.occurrences.count):
        oc = root.occurrences.item(i)
        oc.isLightBulbOn = False
    # Show only the target branch and ensure body visibility for body case.
    if kind == 'component':
        # Turn on only this occurrence (and Fusion will show its bodies as they were)
        ref.isLightBulbOn = True
    else:  # body
        # Show the body’s owning occurrence path
        parent_occ = ref.assemblyContext  # may be None if body is in root
        if parent_occ:
            parent_occ.isLightBulbOn = True
        # Hide all bodies, then show only target body
        for i in range(root.bRepBodies.count):
            b = root.bRepBodies.item(i)
            b.isVisible = False
        ref.isVisible = True

def _export_mesh(design, geometry : adsk.fusion.BRepBody, fullpath, fmt_name):
    # make sure the body is visible
    geometry.isVisible = True
    em = design.exportManager
    fmt = fmt_name.upper()
    _app.log("exporting format: '{}' path: '{}'".format(fmt, fullpath))
    if fmt == 'STL':
        # STL accepts Body/Occurrence/Component
        opts = em.createSTLExportOptions(geometry, fullpath)
        # Optional: opts.meshRefinement = adsk.fusion.MeshRefinementSettings.MeshRefinementMedium
        return em.execute(opts)
    elif fmt == 'OBJ':
        opts = em.createOBJExportOptions(geometry, fullpath)
        return em.execute(opts)
    elif fmt == '3MF':
        opts = em.createC3MFExportOptions(geometry, fullpath)
        return em.execute(opts)
    else:
        return False

def _export_step(design, fullpath):
    em = design.exportManager
    # Export “whole design”; visibility filtering already isolated the target
    opts = em.createSTEPExportOptions(fullpath)  # no geometry arg -> root
    return em.execute(opts)

def _progress_start(title, maximum):
    dlg = _ui.createProgressDialog()
    # message supports %p (percent), %v (current), %m (min), %t (max)
    dlg.show(title, 'Exporting… %v / %m  (%p%)', 0, maximum, 0)
    dlg.isBackgroundTranslucent = False
    dlg.cancelButtonText = 'Cancel'
    return dlg

def _progress_update(note=None):
    _prog.progressValue = _step
    if note:
        #_prog.message = f'Exporting… {_step} / {_prog.maximumValue}\n\n{note}'
        _prog.message = f'Exporting… %v / %m  (%p%)\n\n{note}'
    adsk.doEvents()  # keep UI responsive

def _progress_end(dlg):
    try: dlg.hide()
    except: pass