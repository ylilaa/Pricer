Attribute VB_Name = "functions"

' ------- input first cell, function counts until the last cell of dataset and returns the full range --------'
Public Function count_datarange(first_cell) As Range
    Dim i, j As Long
    i  = first_cell.End(xlDown).Row 
    j = first_cell.End(xlToRight).Column 
    Set count_datarange = Range(first_cell,  cells(i,j))
End Function 

'--------- Stores the fund in sheet into an object -----------'
Public  Function build_fund() As Object
    Dim code_cell as Range
    Set code_cell = find_value("Code","Fund_pricing")    
    Set build_fund = count_datarange(code_cell)
End Function

'------------Finds first instance of a value (1st argument) in a worsheet (second argument), returns its cell range
Public Function find_value(value , sheet) As Range
    Dim R As Range
    Set R = Worksheets(sheet).Cells
    Set find_value = R.Find(value)
End Function

'------------Finds first instance of a column (1st argument) in a worsheet (second argument), returns its number
Public Function find_column(col_name , sheet) As Integer
    find_column = Worksheets(sheet).Cells.Find(what:=col_name).column
End Function


'---------- Stores echeancier in an object--------------'
' Public Function build_echeancier(title) As Array
'     Dim echeancier As Variant
    

'     Set build_echeancier = echeancier 
' End Function 

'----------- Recupere un titre depuis la db -------------'
Public Function build_title(title_code) As Object

    Dim row_num As Integer
    Dim titleRange As Object
    Dim title_data As Object
    Set title_data = CreateObject("Scripting.Dictionary")
    Dim temp_col_num As Integer    
    Set titleRange  = find_value(title_code, "Titles_db")
    row_num = titleRange.Row
    
    title_data.Add "Code", title_code

    temp_col_num = find_column("DESCRIPTION","Titles_db")
    title_data.Add "Description", Worksheets("Titles_db").cells(row_num,temp_col_num).value
    
    temp_col_num = find_column("CODE_FONDS","Titles_db")
    title_data.Add "Code_fonds", Worksheets("Titles_db").cells(row_num,temp_col_num).value
    
    temp_col_num = find_column("QUANTITE","Titles_db")
    title_data.Add "Quantite", Worksheets("Titles_db").cells(row_num,temp_col_num).value
    
    temp_col_num = find_column("EMETTEUR","Titles_db")
    title_data.Add "Emetteur", Worksheets("Titles_db").cells(row_num,temp_col_num).value
    
    temp_col_num = find_column("DATE_EMISSION","Titles_db")
    title_data.Add "Date_emission", Worksheets("Titles_db").cells(row_num,temp_col_num).value
    
    temp_col_num = find_column("DATE_JOUISSANCE","Titles_db")
    title_data.Add "Date_jouissance", Worksheets("Titles_db").cells(row_num,temp_col_num).value
    
    temp_col_num = find_column("DATE_ECHEANCE","Titles_db")
    title_data.Add "Date_echeance", Worksheets("Titles_db").cells(row_num,temp_col_num).value
    
    temp_col_num = find_column("NOMINAL","Titles_db")
    title_data.Add "Nominal", Worksheets("Titles_db").cells(row_num,temp_col_num).value
    
    temp_col_num = find_column("MR","Titles_db")
    title_data.Add "Mr", Worksheets("Titles_db").cells(row_num,temp_col_num).value
    
    temp_col_num = find_column("MR_T","Titles_db")
    title_data.Add "Mr_t", Worksheets("Titles_db").cells(row_num,temp_col_num).value
    
    temp_col_num = find_column("SPREAD","Titles_db")
    title_data.Add "Spread", Worksheets("Titles_db").cells(row_num,temp_col_num).value
    
    temp_col_num = find_column("CATEGORIE","Titles_db")
    title_data.Add "Categorie", Worksheets("Titles_db").cells(row_num,temp_col_num).value
    
    temp_col_num = find_column("PERIODICITE","Titles_db")
    title_data.Add "Periodicite", Worksheets("Titles_db").cells(row_num,temp_col_num).value
    
    temp_col_num = find_column("AMORT","Titles_db")
    title_data.Add "Amort", Worksheets("Titles_db").cells(row_num,temp_col_num).value
    
    Set build_title = title_data
End Function

'---------------Fonction qui construit la colonne des dates de tombee entre une date depart a preciser et une date echeance du titre-----------'
Public Function calcul_date_tombee(title_code,date_depart) As Collection
    Dim title As Object
    Dim date_tombee_rows As new Collection
    
    Dim date_tombee As Date
    date_tombee = date_depart
    
    Set title = CreateObject("Scripting.Dictionary")
    Set title = build_title(title_code)
    
    If (title("Periodicite")="AN") Then
        Do While date_tombee <= title("Date_echeance")
            date_tombee_rows.Add date_tombee
            date_tombee = DateAdd("yyyy",1, date_tombee)
        Loop
         
    ElseIf (title("Periodicite")="SEM") Then 
        Do While date_tombee <= title("Date_echeance")
            date_tombee_rows.Add date_tombee
            date_tombee = DateAdd("q",2, date_tombee)
        Loop
    ElseIf (title("Periodicite")="TRI") Then 
        Do While date_tombee <= title("Date_echeance")
            date_tombee_rows.Add date_tombee
            date_tombee = DateAdd("q",1, date_tombee)
        Loop
    ElseIf (title("Periodicite")="MEN") Then 
        Do While date_tombee <= title("Date_echeance")
            date_tombee_rows.Add date_tombee
            date_tombee = DateAdd("m",1, date_tombee)
        Loop
    End If

    Set calcul_date_tombee = date_tombee_rows

End Function

    '----------Calcul colonne amorti-----------'
Public Function calcul_amorti(title_code,date_depart,nbr_coupons) As Collection
    
    Dim title As Object
    Set title = CreateObject("Scripting.Dictionary")
    Set title = build_title(title_code)

    Dim col_amorti as new Collection
    Dim amort As Double
    amort = title("Nominal") / nbr_coupons
    Dim i As Integer

    If (title("Amort")="FIN") Then
        i = 1
        While i < nbr_coupons - 1
            col_amorti.Add 0
            i = i + 1
        Wend
    col_amorti.Add title("Nominal")
    Else
        i = 1
        While i <= nbr_coupons - 1
            col_amorti.Add amort
            i = i + 1
                        
        Wend
    End If
    Set calcul_amorti = col_amorti
End Function
'----------Store all fund titles in an object------------'
Public Function build_fund(fund_name) As Object

End Function
'----------Clear previous fund from Home sheet-----------' 