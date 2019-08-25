
*! Mathew Bidinlib mbidinlib@poverty-action.org


prog define gradeapplication

	syntax using/,clear

	clear

	qui {
	
		import excel using "`using'" , clear 

		*globals for details entered

		ren _all, lower

		loc d3 	 =   d[3]
		loc d4	=    d[4]
		loc	d5	=    d[5]
		loc d6  =    d[6]
		loc d7  =    d[7]
		loc d8  =    d[8]
		loc d9  =    d[9]

		glo folder 	    `d4'
		glo data 	    `d5'
		glo criteria    `d6'
		glo output      `d7'
		glo select_num  `d8'
		glo savevar     `d9'

		*add a dummy variable for merging
		
		if regexm("$folder/$data",".xlx|.xlsx|.xlsm") {
			import excel using "$folder/$data", first clear
			
			esle if regexm("$folder/$data",".csv") {
				import delimited using "$folder/$data", clear
			}
			else if regexm("$folder/$data",".dta") {
				use "$folder/$data", clear
			}
			
			else {
				di as err "Invalid data format. YOu may  have to ad the extension(.dta,.xlx,.csv)"
				exit 999
			}
		}
		
		
		cap gen  az= _n
		save "${folder}/${data}_checked.dta", replace

		import excel using "$folder/$criteria", clear
		ren _all, lower

		merge 1:1 az using "${folder}/${data}_checked.dta", force

		*Loop to Check the variables specified

		foreach i of varlist d h { 
			
			forval j = 11(10)101 {

			*Generate variables for score if not empty
				loc svar= `i'[`j']
				if "`svar'" !="" {
					loc var_1= `i'[`j']
					
					*Confirm if specified variable exists in the data
					confirm  var `var_1'
					
						* Verify if name of variable is not longer than 20
							loc a1= strlen("`var_1'")
							loc b1= substr("`var_1'",1,23)
							loc vars= cond(`a1'<25, "`var_1'","`b1'")
							gen nscore_`vars' =0
									
							*loop though the criteria
							forval k= 2/8 {
							
								loc pval= cond("`i'"== "d", "c" , "g")
								loc val_1 = `j'+`k'
						     	loc cval = `pval'[`val_1']

							    if "`cval'" != "" {
									loc grade_1= `j'+`k'
									loc  cgrade= `i'[`grade_1']
									replace nscore_`vars' = `cgrade'  if `vars' == `cval'
								}

							 }
							
				 }
				 
			}
			
			
		** String Variables
		forval j = 113(8)137 {

			*Generate variables for score if not empty
				loc svar= `i'[`j']
				if "`svar'" !="" {
					loc var_1= `i'[`j']
					
					*Confirm if specified variable exists in the data
					confirm  var `var_1'
					
						* Verify if name of variable is not longer than 20
							loc a1= strlen("`var_1'")
							loc b1= substr("`var_1'",1,23)
							loc vars= cond(`a1'<25, "`var_1'","`b1'")
							gen nscore_`vars' =0
									
							*loop though the criteria
							forval k= 2/6 {
							
								loc pval= cond("`i'"== "d", "c" , "g")
								loc val_1 = `j'+`k'
						     	loc cval = `pval'[`val_1']

							    if "`cval'" != "" {
									loc grade_1= `j'+`k'
									loc  cgrade= `i'[`grade_1']
									replace nscore_`vars' = `cgrade'  if regexm(`vars',"`cval'")
								}

							 }
							
				 }
				 
			}			
			
			
			
			
			
		}

		*drop uneccesary variables and export
		drop b- az 
		drop if _merge==1
		drop _merge

		export excel using "$folder/${output}.xlsx", sheet(raw) sheetreplace firstrow(varl)

		*export excel using $folder/$output, first replace
		*Total Score
		egen nscore_total= rowtotal(nscore*)

		*Rank
		gsort - nscore_total
		
		*selected
		gen selected= "Yes" if _n<= $select_num
		replace selected= "No" if _n> $select_num


		if strlen("${savevar}")>=1 {
			export excel $savevar nscore* selected if nscore_total>0 & selected== "Yes"  using "$folder/$output.xlsx", sheet(selected) sheetreplace firstrow(varl)
			export excel $savevar nscore* selected if nscore_total>0 & selected== "No"   using "$folder/$output.xlsx", sheet(not_selected) sheetreplace firstrow(varl)
			export excel $savevar nscore* selected if nscore_total<0  using "$folder/$output.xlsx", sheet(dropped) sheetreplace firstrow(varl)

			else if strlen("${savevar}")<1 {
			export excel  if nscore_total>0 & selected== "Yes"  using "$folder/$output.xlsx", sheet(selected) sheetreplace firstrow(varl)
			export excel  if nscore_total>0 & selected== "No"   using "$folder/$output.xlsx", sheet(not_selected) sheetreplace firstrow(varl)
			export excel  if nscore_total<0  using "$folder/$output.xlsx", sheet(dropped) sheetreplace firstrow(varl)

			}
	}

end
