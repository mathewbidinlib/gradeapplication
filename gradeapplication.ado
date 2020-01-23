
**********************************************
**Grading applications and first shortlisting*
** August 2019*******************************
**********************************************
*Purpose
* When 

*! Mathew Bidinlib 
*!mbidinlib@poverty-action.com

cap prog drop gradeapplication

prog define gradeapplication

	syntax using/,clear

	cls

	qui {
	
		import excel using "`using'" , clear 

		*locals for details entered

		ren _all, lower

		loc folder 		  =    d[4]
		loc data		  =    d[5]
		loc	criteria	  =    d[6]
		loc output 		  =    d[7]
		loc select_num    =    d[8]
		loc savevar   	  =    d[9]

		
	*** Giving detail of the program
		noi di "{title: Grading Applications for shorlisting}"
		noi di "Folder"		_column(30) "`folder'" 
		noi di "Number to be selected" _column(30)  "`select_num'"
		noi di "{hline}"
		

		*add a dummy variable for merging
		cap confirm var az
		if !_rc {
			destring az, replace
			replace az= _n
			}

		if _rc {
			gen  az= _n
		}
		
		tempfile criteria_sheet
		save criteria_sheet, replace
		
		if regexm("`folder'/`data'",".xlx|.xlsx|.xlsm") import excel using "`folder'/`data'", first clear
			
		else if regexm("`folder'/`data'",".csv") 		import delimited using "`folder'/`data'", clear
		
		else if regexm("`folder'/`data'",".dta") 		use "`folder'/`data'", clear
			
		else {
			di as err "Invalid data format. You may  have to ad the extension(.dta,.xlx,.csv)"
			exit 999
			}
	
				
		gen az=_n
		tempfile data_file
		save data_file, replace

		use "criteria_sheet", clear
		ren _all, lower

		merge 1:1 az using "data_file", force

		*Loop to Check the variables specified

		foreach i of varlist d h { 
			
			forval j = 11(13)651 {

			*Generate variables for score if not empty
				loc svar= `i'[`j']
				if "`svar'" !="" {
					loc var_1= `i'[`j']
					
					*Confirm if specified variable exists in the data
					confirm  var `var_1'
					
						* Verify if name of variable is not longer than 20
							loc a1= strlen("`var_1'")
							loc b1= substr("`var_1'",1,20)
							loc vars= cond(`a1'<20, "`var_1'","`b1'")
							gen nsc_`vars'`i'`j' =0
									
							*loop though the criteria
							forval k= 2/11 {
							
								loc pval= cond("`i'"== "d", "c" , "g")
								loc val_1 = `j'+`k'
						     	loc cval = `pval'[`val_1']

							    if "`cval'" != "" {
									loc grade_1= `j'+`k'
									loc  cgrade= `i'[`grade_1']
									replace nsc_`vars'`i'`j' = `cgrade'  if `var_1' == `cval'
								}

							 }
							
				 }
				 
			}
			
			
		** String Variables
		forval j = 668(13)707 {

			*Generate variables for score if not empty
				loc svar= `i'[`j']
				if "`svar'" !="" {
					loc var_1= `i'[`j']
					
					*Confirm if specified variable exists in the data
					confirm  var `var_1'
					
						* Verify if name of variable is not longer than 20
							loc a1= strlen("`var_1'")
							loc b1= substr("`var_1'",1,15)
							loc vars= cond(`a1'<25, "`var_1'","`b1'")
							gen nsc_`vars' = 0
									
							*loop though the criteria
							forval k= 2/11 {
							
								loc pval= cond("`i'"== "d", "c" , "g")
								loc val_1 = `j'+`k'
						     	loc cval = `pval'[`val_1']

							    if "`cval'" != "" {
									loc grade_1= `j'+`k'
									loc  cgrade= `i'[`grade_1']
									replace nsc_`vars' = `cgrade'  if regexm(`var_1',"`cval'")
								}

							 }
							
				 }
				 
			}			
			
			
			
			
			
		}

		*drop uneccesary variables and export raw data of applicants
		drop b- az 
		drop if _merge==1
		drop _merge

		
		export excel using "`folder'/`output'.xlsx", sheet(raw) sheetreplace firstrow(varl)

		*export excel using $folder/$output, first replace
		*Total Score
		egen total_score= rowtotal(nsc_*)

		*Rank
		gsort - total_score
		
		egen rank = rank(-total_score)
		gen ranking=floor(rank)
		
		*selected
		if strlen("`select_num'")>=1 {
			gen selected= "Yes" if ranking<= `select_num'
			replace selected= "No" if ranking> `select_num'
		}
		else gen selected="Yes"
		

		loc outfile "`folder'/`output'.xlsx"
		
	

		loc savevars "`savevar' nsc* total_score ranking selected"
		drop if missing(username)
		
		noi di "{title: Exporting Sheets}"
		noi di "Sheet 1" _column(30)  "Selected candidates"
		export excel `savevar' nsc* total_score ranking selected if total_score>0 & selected== "Yes"  using "`outfile'", sheet(selected) cell("B2") sheetreplace firstrow(varl)
		
		*Format Sheet containing selected candidates
		count if total_score>0 & selected== "Yes"
		loc  rows = r(N)
		mata: sheet_format("selected")
		
		
		no di "Sheet 2" _column(30)  "Candidates Not Selected"
		cap export excel `savevar' nsc* total_score ranking selected if total_score>0 & selected== "No"   using "`outfile'", sheet(not_selected) cell("B2") sheetreplace firstrow(varl)
		
		*format sheet containing not selected candidates
		count if total_score>0 & selected== "No" 
		loc  rows = r(N)
		mata: sheet_format("not_selected")
		
		noi di "Sheet 3" _column(30)  "Candidates Dropped"
		export excel `savevar' nsc* total_score ranking selected if total_score<0  using "`outfile'", sheet(dropped) cell("B2") sheetreplace firstrow(varl)
		
		*Format sheet of dropped candidates
		count if total_score<0
		loc  rows = r(N)
		mata: sheet_format("dropped")


  }
  
  noi di "Selection completed"
  noi di "{hline}"

end



* Formart exported sheets

mata:
mata clear
void sheet_format(string scalar sheet) {
	filename 		= st_local("outfile")
	vars			= st_local("savevars")
	rows			= strtoreal(st_local("rows"))

	class xl scalar b

	b.load_book(filename)
	b.set_sheet(sheet)
	b.set_sheet_gridlines(sheet, "off")
	
	dat= st_data(.,vars)

	b.set_border((2, rows + 2), (2, cols(dat) + 1), "thin")
	b.set_top_border((2, 2),  (2, cols(dat) + 1), "thick")
	b.set_bottom_border((2, 2),  (2, cols(dat) + 1), "thick")
	b.set_left_border((2, rows + 2), (2, 2), "thick")
	b.set_right_border((2, rows + 2), (cols(dat) + 1, cols(dat) + 1), "thick")
	b.set_bottom_border((rows + 2, rows + 2), (2, cols(dat) + 1), "thick")

}
end












