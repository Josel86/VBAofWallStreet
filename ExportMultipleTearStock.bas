Attribute VB_Name = "Módulo1"
Sub MultipleYearStock()

    ' Inicia el recorrido de cada hoja
    For Each ws In Worksheets
                
        ' Se crea una variable con el nombre de la hoja
        Dim WorksheetName As String

        ' Se determina el ultimo renglon de la hoja
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Se determine la ultima columna de la hoja
        LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column

        ' Indica el nombre del Ticker
        Dim Ticker_Name As String
      
        ' Coloca la variable Inical de la suma de volumenes por Ticker
        Dim Stock_Vol As Double
        Stock_Vol = 0

        ' Coloca el valor cuando abre la accion
        Dim Stock_Open As Double
        Stock_Open = 0

        ' Coloca el valor cuando cierra la accion
        Dim Stock_Close As Double
        Stock_Close = 0

        ' Bandera que indica el inicio de un periodo
        Dim Stock_Open_Flag As Boolean
        Stock_Open_Flag = True

        ' Indica el valor del porcentaje de cambio anual
        Dim Stock_Percent_Change As Double
        Stock_Percent_Change = 0

        ' Indica el valor del cambio anual
        Dim Stock_Year_Change As Double
        Stock_Year_Change = 0

        ' Indices de la tabla resumen
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2

        'Indica los Tickers mas y menos grandes
        Dim Increase_Ticker As String
        Dim Decrease_Ticker As String
        Dim Volume_Ticker As String
        
        'Indica los vaores de los Tickers mas y menos grandes
        Dim Increase_Ticker_Value As Double
        Dim Decrease_Ticker_Value As Double
        Dim Volume_Ticker_Value As Double
        Increase_Ticker_Value = 0
        Decrease_Ticker_Value = 0
        Volume_Ticker_Value = 0

        ' Se ingresan los encabezados de las tablas
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        ' Se recorren cada uno de los registros apartir del renglon 2
        For i = 2 To LastRow
            If Stock_Open_Flag = True Then
                Stock_Open = ws.Cells(i, 3).Value
                Stock_Open_Flag = False
            End If
            
            ' Se valida que el valor anterior cambia con el fin de contabilizar todos los registros de un Ticker, siempre y cuando esten ordenados
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                ' Asgina el valor del cambio
                Stock_Close = ws.Cells(i, 6).Value
                
                ' Colocar el nombre de el Ticker
                Ticker_Name = ws.Cells(i, 1).Value

                ' Agrega el valor total del Stocker
                Stock_Vol = Stock_Vol + ws.Cells(i, 7).Value

                ' Calcula el cambio anual
                Stock_Year_Change = Stock_Close - Stock_Open
               
                ' Calcula el porcentaje de cambio anual
                If Stock_Open = 0 Then
                    Stock_Percent_Change = 1
                Else
                    Stock_Percent_Change = (Stock_Year_Change / Stock_Open)
                End If
                
                ' Escribe el valor del ticker
                ws.Range("I" & Summary_Table_Row).Value = Ticker_Name

                ' Escribe el valor del cambio de año
                ws.Range("J" & Summary_Table_Row).Value = Stock_Year_Change
                ws.Range("J" & Summary_Table_Row).NumberFormat = "0.000000000000000"
                If (Stock_Year_Change >= 0) Then
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
                
                ' Escribe el valor porcentaje de cambio
                ws.Range("K" & Summary_Table_Row).Value = Stock_Percent_Change
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                ' Escribe el valor total del Stocker
                ws.Range("L" & Summary_Table_Row).Value = Stock_Vol
 
                ' Actualiza el contador para escribir en el siguiente renglon
                Summary_Table_Row = Summary_Table_Row + 1
                
                ' Valida si es el Stock Percent mas grande
                If (Stock_Percent_Change > Increase_Ticker_Value) Then
                    Increase_Ticker_Value = Stock_Percent_Change
                    Increase_Ticker = Ticker_Name
                End If
                
                ' Valida si es el Stock Percent mas pequeño
                If (Stock_Percent_Change < Decrease_Ticker_Value) Then
                    Decrease_Ticker_Value = Stock_Percent_Change
                    Decrease_Ticker = Ticker_Name
                End If
                
                ' Valida si es el Volume mas grande
                If (Stock_Vol > Volume_Ticker_Value) Then
                    Volume_Ticker_Value = Stock_Vol
                    Volume_Ticker = Ticker_Name
                End If
                
                ' Re-establece el contados
                Stock_Vol = 0
                Stock_Open = 0
                Stock_Close = 0
                Stock_Year_Change = 0
                Stock_Percent_Change = 0
                Stock_Open_Flag = True
            Else
                ' Suma el valor del Stocker
                 Stock_Vol = Stock_Vol + ws.Cells(i, 7).Value
            End If
        Next i
        
        'Se imprimen el resumen de la ultima tabla
        ws.Range("P2").Value = Increase_Ticker
        ws.Range("P3").Value = Decrease_Ticker
        ws.Range("P4").Value = Volume_Ticker
        
        ws.Range("Q2").Value = Increase_Ticker_Value
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").Value = Decrease_Ticker_Value
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("Q4").Value = Volume_Ticker_Value
        
    Next ws
End Sub
    
