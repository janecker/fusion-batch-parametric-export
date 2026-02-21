# Batch Parametric export for Autodesk Fusion

This scripts allows to batch export your Fusion design with different parameter combinations. Export possible in mesh formats and STEP.

Adapation of the original script https://github.com/andrewtch/batch-parametric-export for my specific usecase.

# Usage

Add the script folder to Fusion, run the script, then:

- Add parameter 'output_filename' with parameters mentioned below, e.g.: '{selection_set}_{param_values}/{project_name}_{param_values}_{selection_set}_{body_name}'
- Add parameter 'output_formats' e.g.: 'STL;STEP'
- Make sure the parameters that should be considered in the export are favourites and add a comment with the parameter combinations as semicolon list, e.g: 4.1;4.2;4.3
- Create selection sets of the bodies you want to export, name them Static_XYZ or Parametric_XYZ. Static bodies will be exported as is, parametric bodies will be exported with all parameter combinations.
- Run the script, select an output folder and execute

# 'output_filename' Options:

- project_name: Name of the document including version number you are exporting
- param_values: list of all the parameter values you are exporting seprated by _
- selection_set: Name of the selection set, without the Static_ or Parametric_ part (e.g. XYZ for the example above)
- body_name: Name of the body

# Implementation details

- only non-computable parameters are supported and displayed in the interface (2mm works; 19/param3*3.14 does not)
- text parameters work (Fusion Sep-2025+).  "'dog';'cat';'bob';'sue'". 
- absolute path is not required, relative (~/Documents) may fail


# Todo

- lots of checking etc. probably missing
