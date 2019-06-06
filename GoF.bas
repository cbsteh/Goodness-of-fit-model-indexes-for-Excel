Attribute VB_Name = "GoF"
' # GoF
' Several model goodness-of-fit (GoF) indexes for Excel
'
'
' Author: Christopher Teh Boon Sung, Uni. Putra Malaysia
'
' Contact: christeh@yahoo.com; www.christopherteh.com
'
' Initial Release: June 6, 2019
'
'
' MIT -licensed:
' *  Free to use, copy, share, and modify
' *  Give credit to the developer somewhere in your software code or documentation
'
'
' List of GoF indexes (and the names of their functions in brackets):
' 1.  Mean Absolute Error (`fit_mae`)
' 1.  Normalized Mean Absolute Error (`fit_nmae`)
' 1.  Mean Bias Error (P-O) (`fit_mbe`)
' 1.  Normalized Mean Bias Error (P-O) (`fit_nmbe`)
' 1.  Root Mean Square Error (`fit_rmse`)
' 1.  Original Index of Agreement (`fit_d`)
' 1.  New (Refined) Index of Agreement (`fit_dr`)
' 1.  RMSE to Standard Deviation Ratio (`fit_rsr`)
' 1.  Nash-Sutcliffe Efficiency (`fit_nse`)
' 1.  Normalized mean square error (`fit_nmse`)
' 1.  Fractional bias (`fit_fb`)
' 1.  Coefficient of Efficiency (`fit_coe`)
' 1.  Revised Mielke Index (`fit_mielke`)
' 1.  Persistence Index (`fit_pi`)
' 1.  Akaike Information Criterion (AIC) (`fit_aic`)
' 1.  Bayesian Information Criterion (BIC) (`fit_bic`)
'
'
' Note:
' *  All indexes will ignore cells that are blank (empty), hidden, or contain `#N/A` error
' *  Missing values in cells should be left blank or use the function `NA()` to indicate an error value in that cell
'
'
' Installation:
' 1. Open the Visual Basic Editor in Excel (via the Developer tab)
' 1. Insert the file (`Gof.bas`) as one of the modules in your workbook.
'
'
' Usage:
' * All GoF functions start with `fit_<<name>>` where `<<name>>` is the abbreviated name of the GoF index. For instance, the mean bias error (MBE) index function is `fit_mbe`, and the normalized mean absolute error (NMAE) function is `fit_nmae`. See the GoF module for the other functions.
' * To use the MBE function, type in `=fit_mbe(A1:A10, B1:B10)`, where `A1:A10` is the range of cells containing the observed (measured) values and `B1::B10` the estimated (predicted) values. Other GoF functions are used in the same way, except for AIC and BIC functions.
' * To use the AIC function, type in `=fit_aic(A1:A10, B1:B10, 3, True)` where `A1:A10` and `B1:B10` contain the observed and estimated values, respectively; the third argument (value `3`) is the number of model parameters plus one (e.g., simple linear regression equation y = mx + c has 3 model parameters: m, c, and plus one); and the last parameter is True (by default) for second-order AIC. Set to False for first order AIC (use for large samples).
' * The BIC function is used in the same way as the AIC function, except the BIC function is `fit_bic` and it has no fourth parameter, e.g., `=fit_bic(A1:A10, B1:B10, 3)`.
'
  

Private Function FillInValues(obs As Range, est As Range, co As Variant, cp As Variant)
' ** Internal use **
' Collect valid pairwise values (obs, est). Ignores any cells that contain blanks and errors or cellt that are hidden.
' Returns 0 if pairwise values have been collected and are valid. A return of any other value indicates one or more cells are invalid.
'
   Dim sz As Integer
   sz = obs.Count
   If sz <> est.Count Or sz < 2 Then
      FillInValues = CVErr(xlErrValue)    ' unequal or insufficient size
      Exit Function
   End If

   Dim i As Integer, n As Integer
   n = 0
   For i = 1 To sz
      ' fill in values, but skip blanks, error cell values, or non-numeric cells
      If CheckValue(obs.Cells(i)) = 1 And CheckValue(est.Cells(i)) = 1 Then
         ReDim Preserve co(n)
         ReDim Preserve cp(n)
         co(n) = obs.Cells(i)
         cp(n) = est.Cells(i)
         n = n + 1
      End If
   Next i
   
   If n < 2 Then
      FillInValues = CVErr(xlErrValue)    ' need at least two pairs of values
   Else
      FillInValues = 0     ' no error
   End If

End Function

Private Function CheckValue(r As Range)
' ** Internal use **
' Checks if a cell is blank, hidden, or contain the error #N/A
' Returns 1 if cell has a valid number, -1 if it is blank, hidden, or has the error #N/A, else -2 if a cell has invalid number (e.g., contains a text)
'

   If IsEmpty(r.Value) Then
      CheckValue = -1    ' ok, skip
   ElseIf IsError(r.Value) Then
      If r.Value = CVErr(xlErrNA) Then
         CheckValue = -1 ' ok, skip
      End If
   ElseIf r.EntireRow.Hidden Or r.EntireColumn.Hidden Then
      CheckValue = -1   ' ok, skip
   ElseIf IsNumeric(r.Value) Then
      CheckValue = 1    ' ok, use
   Else
      CheckValue = -2   ' not ok
   End If
      
End Function

Private Function Average(ar() As Variant)
' ** Internal use **
' Returns the average of an array
'
   Dim i As Integer
   Dim sum As Double
   sum = 0#
   For i = LBound(ar) To UBound(ar)
      sum = sum + ar(i)
   Next i
   Average = sum / (UBound(ar) - LBound(ar) + 1)

End Function

Private Function Correlation(x() As Variant, y() As Variant)
' ** Internal use **
' Returns the correlation coefficient of an array
'
   Dim sum As Double, meanx As Double, meany As Double
   sum = 0#
   meanx = Average(x)
   meany = Average(y)
   Dim i As Integer
   For i = LBound(x) To UBound(x)
      sum = sum + ((x(i) - meanx) * (y(i) - meany))
   Next i
   sum = sum / (UBound(x) - LBound(x))
   Correlation = sum / (StdDev(x) * StdDev(y))

End Function

Private Function StdDev(ar() As Variant)
' ** Internal use **
' Returns the standard deviation of an array
'
   Dim i As Integer
   Dim sum As Double, mean As Double
   mean = Average(ar)
   sum = 0#
   For i = LBound(ar) To UBound(ar)
      sum = sum + (ar(i) - mean) ^ 2
   Next i
   StdDev = (sum / (UBound(ar) - LBound(ar))) ^ 0.5

End Function

Function fit_mae(obs As Range, est As Range)
' Mean Absolute Error |P-O|
' Parameters: obs = observed values; est = estimated (predicted) values
' Range: 0 to +INF
' Best fit = 0, large +ve = large errors
'
   Dim co() As Variant, cp() As Variant
   fit_mae = FillInValues(obs, est, co, cp)
   If fit_mae <> 0 Then
      Exit Function
   End If
   
   Dim n1 As Double
   n1 = 0#
   Dim i As Integer
   For i = LBound(co) To UBound(co)
      n1 = n1 + Abs(cp(i) - co(i))
   Next i
   fit_mae = n1 / (UBound(co) - LBound(co) + 1)

End Function

Function fit_nmae(obs As Range, est As Range)
' Normalized Mean Absolute Error |P-O| / (mean O)
' Parameters: obs = observed values; est = estimated (predicted) values
' Range: 0 to +INF
' Best fit = 0, large +ve = large errors
'
   Dim co() As Variant, cp() As Variant
   fit_nmae = FillInValues(obs, est, co, cp)
   If fit_nmae <> 0 Then
      Exit Function
   End If
   
   Dim n1 As Double, n2 As Double
   n1 = 0#
   n2 = 0#
   Dim i As Integer
   For i = LBound(co) To UBound(co)
      n1 = n1 + Abs(cp(i) - co(i))
      n2 = n2 + co(i)
   Next i
   fit_nmae = n1 / n2

End Function

Function fit_mbe(obs As Range, est As Range)
' Mean Bias Error (P-O)
' Parameters: obs = observed values; est = estimated (predicted) values
' Range: -INF to +INF
' Best fit = 0, large +ve = overestimate, large -ve = underestimate
' 0.10 for very good; 0.10 - 0.15 for good and 0.15 - 0.25 for satisfactory ratings
'
   Dim co() As Variant, cp() As Variant
   fit_mbe = FillInValues(obs, est, co, cp)
   If fit_mbe <> 0 Then
      Exit Function
   End If
   
   Dim n1 As Double
   n1 = 0#
   Dim i As Integer
   For i = LBound(co) To UBound(co)
      n1 = n1 + (cp(i) - co(i))
   Next i
   fit_mbe = n1 / (UBound(co) - LBound(co) + 1)

End Function

Function fit_nmbe(obs As Range, est As Range)
' Normalized Mean Bias Error (P-O) / (mean O)
' Parameters: obs = observed values; est = estimated (predicted) values
' Range: -INF to +INF
' Best fit = 0, large +ve = overestimate, large -ve = underestimate
'
   Dim co() As Variant, cp() As Variant
   fit_nmbe = FillInValues(obs, est, co, cp)
   If fit_nmbe <> 0 Then
      Exit Function
   End If
   
   Dim n1 As Double, n2 As Double
   n1 = 0#
   n2 = 0#
   Dim i As Integer
   For i = LBound(co) To UBound(co)
      n1 = n1 + (cp(i) - co(i))
      n2 = n2 + co(i)
   Next i
   fit_nmbe = n1 / n2

End Function

Function fit_rmse(obs As Range, est As Range)
' Root Mean Square Error (P-O)
' Parameters: obs = observed values; est = estimated (predicted) values
' Range: 0 to +INF
' Best fit = 0, large +ve = large errors
'
   Dim co() As Variant, cp() As Variant
   fit_rmse = FillInValues(obs, est, co, cp)
   If fit_rmse <> 0 Then
      Exit Function
   End If
   
   Dim n1 As Double
   n1 = 0#
   Dim i As Integer
   For i = LBound(co) To UBound(co)
      n1 = n1 + (cp(i) - co(i)) ^ 2
   Next i
   fit_rmse = (n1 / (UBound(co) - LBound(co) + 1)) ^ 0.5

End Function

Function fit_d(obs As Range, est As Range)
' (Original) Index of Agreement, d
' Parameters: obs = observed values; est = estimated (predicted) values
' Range: 0 to 1
' Best fit = 1, Worst fit = 0
' Ref: Willmott, C. J. (1981). On the validation of models. Physical Geography, 2, 184194.
'
   Dim co() As Variant, cp() As Variant
   fit_d = FillInValues(obs, est, co, cp)
   If fit_d <> 0 Then
      Exit Function
   End If
   
   Dim mean_co As Double, n1 As Double, n2 As Double
   mean_co = Average(co)
   n1 = 0#
   n2 = 0#
   Dim i As Integer
   For i = LBound(co) To UBound(co)
      n1 = n1 + Abs(cp(i) - co(i))
      n2 = n2 + (Abs(cp(i) - mean_co) + Abs(co(i) - mean_co))
   Next i
   fit_d = 1 - n1 / n2

End Function

Function fit_dr(obs As Range, est As Range)
' New (Refined) Index of Agreement, dr
' Parameters: obs = observed values; est = estimated (predicted) values
' Range: -1 to 1
' Best fit = 1, Worst fit = -1 (or lack of data/variation)
' Ref: Willmott, C. J., Robeson, S. M., and Matsuura, K. (2012). A refined index of model performance, International Journal of Climatolology, 32: 2088-2094.
' Ref: Willmott, C. J. (1981). On the validation of models. Physical Geography, 2, 184194.
'
   Dim co() As Variant, cp() As Variant
   fit_dr = FillInValues(obs, est, co, cp)
   If fit_dr <> 0 Then
      Exit Function
   End If
   
   Dim mean_co As Double, n1 As Double, n2 As Double
   mean_co = Average(co)
   n1 = 0#
   n2 = 0#
   Dim i As Integer
   For i = LBound(co) To UBound(co)
      n1 = n1 + Abs(cp(i) - co(i))
      n2 = n2 + Abs(co(i) - mean_co)
   Next i

   n2 = 2 * n2
   If n1 <= n2 Then
      fit_dr = 1 - n1 / n2
   Else
      fit_dr = n2 / n1 - 1
   End If

End Function

Function fit_rsr(obs As Range, est As Range)
' RMSE to Standard Deviation Ratio
' Parameters: obs = observed values; est = estimated (predicted) values
' Range: 0 to 1
' Best fit = 0, < 0.50 for very good; 0.50 - 0.60 for good and 0.60 - 0.70 for satisfactory ratings
' Ref: Moriasi, D. N., Arnold, J. G., Van Liew, M. W., Bingner, R. L., Harmel, R. D., and Veith, T. L. (2007). Model evaluation guidelines for systematic quantification of accuracy in watershed simulations. Transactions of the ASABE, 50:885-900.
'
   Dim co() As Variant, cp() As Variant
   fit_rsr = FillInValues(obs, est, co, cp)
   If fit_rsr <> 0 Then
      Exit Function
   End If
   
   Dim mean_co As Double, n1 As Double, n2 As Double
   mean_co = Average(co)
   n1 = 0#
   n2 = 0#
   Dim i As Integer
   For i = LBound(co) To UBound(co)
      n1 = n1 + (cp(i) - co(i)) ^ 2
      n2 = n2 + (co(i) - mean_co) ^ 2
   Next i
   fit_rsr = n1 ^ 0.5 / n2 ^ 0.5

End Function

Function fit_nse(obs As Range, est As Range)
' Nash-Sutcliffe Efficiency
' Parameters: obs = observed values; est = estimated (predicted) values
' Range: -INF to 1
' Best fit = 1, >0.75 for very good; 0.75-0.65 for good and 0.65-0.50 for satisfactory ratings
' Ref: Nash, J. E., and Sutcliffe, J. V. (1970). River flow forecasting through conceptual models part I  A discussion of principles. Journal of Hydrology, 10: 282290.
'
   Dim co() As Variant, cp() As Variant
   fit_nse = FillInValues(obs, est, co, cp)
   If fit_nse <> 0 Then
      Exit Function
   End If
   
   Dim mean_co As Double, n1 As Double, n2 As Double
   mean_co = Average(co)
   n1 = 0#
   n2 = 0#
   Dim i As Integer
   For i = LBound(co) To UBound(co)
      n1 = n1 + (cp(i) - co(i)) ^ 2
      n2 = n2 + (co(i) - mean_co) ^ 2
   Next i
   fit_nse = 1 - n1 / n2

End Function

Function fit_nmse(obs As Range, est As Range)
' Normalized mean square error
' Parameters: obs = observed values; est = estimated (predicted) values
' Range: -INF to +INF
' Best fit = 0, between -0.5 and +0.5 acceptable
'
   Dim co() As Variant, cp() As Variant
   fit_nmse = FillInValues(obs, est, co, cp)
   If fit_nmse <> 0 Then
      Exit Function
   End If
   
   Dim mean_co As Double, mean_cp As Double, n1 As Double
   mean_co = Average(co)
   mean_cp = Average(cp)
   n1 = 0#
   Dim i As Integer
   For i = LBound(co) To UBound(co)
      n1 = n1 + (cp(i) - co(i)) ^ 2
   Next i
   fit_nmse = (n1 / (UBound(co) - LBound(co) + 1)) / (mean_co * mean_cp)

End Function

Function fit_fb(obs As Range, est As Range)
' Fractional bias
' Parameters: obs = observed values; est = estimated (predicted) values
' Range: -INF to +INF
' Best fit = 0, between -0.5 and +0.5 acceptable
'
   Dim co() As Variant, cp() As Variant
   fit_fb = FillInValues(obs, est, co, cp)
   If fit_fb <> 0 Then
      Exit Function
   End If
   
   Dim n1 As Double
   n1 = 0#
   Dim i As Integer
   For i = LBound(co) To UBound(co)
      n1 = n1 + (cp(i) - co(i)) / (0.5 * (cp(i) + co(i)))
   Next i
   fit_fb = n1 / (UBound(co) - LBound(co) + 1)

End Function

Function fit_coe(obs As Range, est As Range)
' Coefficient of Efficiency
' Parameters: obs = observed values; est = estimated (predicted) values
' Range: -INF to 1
' Best fit = 1
'
   Dim co() As Variant, cp() As Variant
   fit_coe = FillInValues(obs, est, co, cp)
   If fit_coe <> 0 Then
      Exit Function
   End If
   
   Dim mean_co As Double, n1 As Double, n2 As Double
   mean_co = Average(co)
   n1 = 0#
   n2 = 0#
   Dim i As Integer
   For i = LBound(co) To UBound(co)
      n1 = n1 + Abs(cp(i) - co(i))
      n2 = n2 + Abs(co(i) - mean_co)
   Next i
   fit_coe = 1 - n1 / n2

End Function

Function fit_mielke(obs As Range, est As Range)
' Revised Mielke Index (Reduction in r due to model errors)
' Parameters: obs = observed values; est = estimated (predicted) values
' Range: -1 to 1
' Best fit = 1
' Ref: Duveiller, G., Fasbender, D., and Meroni, M. (2016). Revisiting the concept of a symmetric index of agreement for continuous datasets. Scientific Reports, 6(19401): 1-14.
' Ref: Mielke, P. (1984). Meteorological applications of permutation techniques based on distance functions. In Krishnaiah, P. & Sen, P. (eds.). Handbook of Statistics Vol. 4, 813830 (Elsevier, Amsterdam, The Netherlands.
'
   Dim co() As Variant, cp() As Variant
   fit_mielke = FillInValues(obs, est, co, cp)
   If fit_miekle <> 0 Then
      Exit Function
   End If
   
   Dim sdx As Double, sdy As Double, meanx As Double, meany As Double, r As Double
   sdx = StdDev(co)
   sdy = StdDev(cp)
   meanx = Average(co)
   meany = Average(cp)
   r = Correlation(co, cp)
   Dim alpha As Double
   alpha = (sdx / sdy) + (sdy / sdx) + ((meanx - meany) ^ 2) / (sdx * sdy)
   fit_mielke = 2 / alpha * r

End Function

Function fit_pi(obs As Range, est As Range)
' Persistence Index
' Parameters: obs = observed values; est = estimated (predicted) values
' Range: -INF to 1
' Best fit = 1, > 0 satisfactory, <= 0 poor
' Ref: Gupta, H. V., Sorooshian, S., and Yapo, P. O. (1998). Toward improved calibration of hydrologic models: multiple and non-commensurable measures of information. Water Resources Research, 34: 751763.
'
   Dim co() As Variant, cp() As Variant
   fit_pi = FillInValues(obs, est, co, cp)
   If fit_pi <> 0 Then
      Exit Function
   End If
   
   Dim n1 As Double, n2 As Double
   n1 = 0#
   n2 = 0#
   Dim i As Integer
   For i = LBound(co) To UBound(co) - 1
      n1 = n1 + (cp(i + 1) - co(i + 1)) ^ 2
      n2 = n2 + (co(i + 1) - co(i)) ^ 2
   Next i
   fit_pi = 1 - n1 / n2

End Function

Function fit_aic(obs As Range, est As Range, k As Integer, Optional bOrder2 = True)
' Akaikes Information Criterion (AIC)
' Parameters: obs = observed values; est = estimated (predicted) values;
' Parameters: k = no. of model parameters plus one; bOrder2 = True for second-order AIC, else False for first-order AIC
' Example: a simple linear regression equation, y = mx + c, has 3 parameters (m and c parameters + 1)
' Returns the second-order AIC if bOrder2 is True (default), else first-order.
' Note: by itself, AIC values have no meaning. AIC is meant to be used to
' compare between models, where the best model is one with the lowest AIC value.
' Ref: Burnham, K. P., and Anderson, D. R. (2002). Model Selection and Multimodel Inference: A practical information-theoretic approach (2nd ed.). Springer-Verlag, NY.
' Ref: Burnham, K. P., and Anderson, D. R. (2004). Multimodel inference: understanding AIC and BIC in Model Selection. Sociological Methods & Research, 33: 261304.
'
   Dim co() As Variant, cp() As Variant
   aic = FillInValues(obs, est, co, cp)
   If aic <> 0 Then
      Exit Function
   End If
   
   Dim rss As Double
   rss = 0#
   Dim i As Integer
   For i = LBound(co) To UBound(co)
      rss = rss + (cp(i) - co(i)) ^ 2
   Next i
   
   Dim n As Long
   n = UBound(co) - LBound(co) + 1
   Dim mle As Double
   mle = rss / n
   fit_aic = -2 * Log(mle) + 2 * k   ' first-order AIC
   
   If bOrder2 Then
      fit_aic = fit_aic + 2 * k * (k + 1) / (n - k - 1)    ' second-order AIC
   End If

End Function

Function fit_bic(obs As Range, est As Range, k As Integer)
' Bayesian information criterion (BIC)
' Parameters: obs = observed values; est = estimated (predicted) values;
' Parameters: k = no. of model parameters plus one
' Example: a simple linear regression equation, y = mx + c, has 3 parameters (m and c parameters + 1)
' Note: by itself, BIC values have no meaning. BIC is meant to be used to
' compare between models, where the best model is one with the lowest BIC value.
' Ref: Burnham, K. P., and Anderson, D. R. (2002). Model Selection and Multimodel Inference: A practical information-theoretic approach (2nd ed.). Springer-Verlag, NY.
' Ref: Burnham, K. P., and Anderson, D. R. (2004). Multimodel inference: understanding AIC and BIC in Model Selection. Sociological Methods & Research, 33: 261304.
'
   Dim co() As Variant, cp() As Variant
   aic = FillInValues(obs, est, co, cp)
   If aic <> 0 Then
      Exit Function
   End If
   
   Dim rss As Double
   rss = 0#
   Dim i As Integer
   For i = LBound(co) To UBound(co)
      rss = rss + (cp(i) - co(i)) ^ 2
   Next i
   
   Dim n As Long
   n = UBound(co) - LBound(co) + 1
   Dim mle As Double
   mle = rss / n
   fit_bic = -2 * Log(mle) + k * Log(n)

End Function
