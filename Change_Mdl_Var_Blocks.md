Add capability for two new recipe instructions, `Input_VarBlock` and `Formula_VarBlock` that refesh indexed blocks of Scenario Model (aka mdl) variables that have the same name with the exception of trailing suffixes `_1`, `_2` etc. from 1 to n where integer n is specified as a JSON ExcelSteps, params column string, {nBlockVars:n}.
* n is an integer. The JSON specification of n is in the recipe row containing the instruction.
* The `Input_VarBlock` ExcelSteps recipe instruction (in the recipe's Step column) simply inserts and/or refreshes the Number Format of a block of input variables (no formula and formula formatting)
* The `Formula_VarBlock` instruction also refreshes the variables' formulas and formats the variable's mdl scenario cells as calculated (same formatting as with current Col_Insert recipe instructions)
* variable blocks are required to be contiguous in the mdl --the variables cannot have blank row gaps or other variables interposed among the individual block variables
* In ExcelSteps, the recipe will list just the "base" variable name for a block variable. The name must have a `_xx` suffix such as  the example, `inputvar_xx`
* If a `Formula_VarBlock` variable's formula refers to other block variables, the recipe Column D formula should show those variables with `_xx` suffix (e.g. `calculatedvar_xx` formula example `=inputvar_xx * 2`). mdl.Refresh needs to resolve this as `=inputvar_1 * 2` for `calculatedvar_1` etc.
* The Input_VarBlock and Formula_VarBlock recipe instructions are only applicable for non-default, "Lite" Scenario Models which use the ExcelSteps recipe to specify instructions about for Number Format and formula -- as opposed to default models that have additional template columns to specify variables' Number Format and Formula.
* There are three additional use cases that `mdl.Refresh` can encounter
	* If the Scenario Model has not previously been refreshed, it may contain the `inputvar_xx` base variable in `mdl.colrngVarNames` to locate the block within the model. When the mdl is refreshed, the `_xx` placeholder row gets replaced by the block variables such as `inputvar_1`, `inputvar_2` and `inputvar_3` if n = 3.
	* If the first refresh of the mdl is to a template, the model may already contain m input or calculated block variables where m is an upper, allowed block size for the application. In this case, n must be specified in the recipe to be less than or equal to m (Fail early if not), and mdl.Refresh should delete the variable rows for greater than n.
	* If the model has been previously refreshed, it may already contain n variables in the block that can be located by searching for the first variable such as `inputvar_1`. In this case, mdl.Refresh should just refresh Number Format and formula (but should not delete and re-add `inputvar_xx` variables in case they already contain data). If the counted number of variables in the existing block is different than the specified n, the block rows should be deleted/replaced with newly-inserted rows. This could occur if the application modifies n between refreshes


Suggest creating a new MdlBlockVar class to manage attributes and methods related to block variable refresh (or would it be simpler to just add block-related attributes to existing mdlRow.cls??). In addition to the nBlockVars attribute, class can store the base variable name and its stem (e.g. `inputvar_xx` base name has stem `inputvar` used to construct indexed mdl names like `inputvar_1` etc.). Class can have a wrapper customized for block variables to manage instancing the existing `mdlRow` class to refresh. This wrapper is needed since refreshing a block variable involves looping to replace `_xx` with the block variable index for both the variable name and/or formula in the case of `Formula_VarBlock` recipe instruction

Possible flow when refresh encounters instruction 
1. Set `varnameStem` as attribute 
2. Search model for base \_xx name or \_1 variable name
3. If base name found insert row for vars and delete base row
4. If \_1 name count 
5. If count matches n move to blockvar refresh
6. If count not match and count > n, delete from bottom of block up until n indexed variables remain
7. if count not match and count < n, delete all and start over adding n indexed variable rows
8. Refresh formulas (special logic to set mdlRow variable name and indexed formula from `varnameStem` attribute and index)

References:
mdlScenario.cls and its .Refresh method
mdlRow.cls
[[Example_Mdl_BlockVars]] showing examples of block variables in a Lite model
[[Example_Mdl_BlockVars_Recipe]] showing a recipe mockup for block variables