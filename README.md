# GoF
Several goodness-of-fit (GoF) model indexes for Excel


Author: Christopher Teh Boon Sung, Uni. Putra Malaysia

Contact: christeh@yahoo.com; www.christopherteh.com

Initial Release: June 6, 2019
Updated: June 30, 2019


MIT -licensed:
*  Free to use, copy, share, and modify
*  Give credit to the developer somewhere in your software code or documentation


List of GoF indexes (and the names of their functions in brackets):
1.  Mean Absolute Error (`fit_mae`)
1.  Normalized Mean Absolute Error (`fit_nmae`)
1.  Mean Bias Error (P-O) (`fit_mbe`)
1.  Normalized Mean Bias Error (P-O) (`fit_nmbe`)
1.  Root Mean Square Error (`fit_rmse`)
1.  Original Index of Agreement (`fit_d`)
1.  New (Refined) Index of Agreement (`fit_dr`)
1.  RMSE to Standard Deviation Ratio (`fit_rsr`)
1.  Nash-Sutcliffe Efficiency (`fit_nse`)
1.  Normalized mean square error (`fit_nmse`)
1.  Fractional bias (`fit_fb`)
1.  Coefficient of Efficiency (`fit_coe`)
1.  Revised Mielke Index (`fit_mielke`)
1.  Persistence Index (`fit_pi`)
1.  Akaike Information Criterion (AIC) (`fit_aic`)
1.  Bayesian Information Criterion (BIC) (`fit_bic`)
1.  Theil's U2 Coefficient of Inequality (UII) (`fit_theilu2`)
1.  Mean Absolute Percentage Error (MAPE) (`fit_mape`)
1.  Median Absolute Percentage Error (MAPE) (`fit_mdape`)


Note:
*  All indexes will ignore cells that are blank (empty), hidden, or contain `#N/A` error
*  Missing values in cells should be left blank or use the function `NA()` to indicate an error value in that cell


Installation:
1. Open the Visual Basic Editor in Excel (via the Developer tab)
1. Insert the file (`Gof.bas`) as one of the modules in your workbook.


Usage:
* All GoF functions start with `fit_<<name>>` where `<<name>>` is the abbreviated name of the GoF index. For instance, the mean bias error (MBE) index function is `fit_mbe`, and the normalized mean absolute error (NMAE) function is `fit_nmae`. See the GoF module for the other functions.
* To use the MBE function, type in `=fit_mbe(A1:A10, B1:B10)`, where `A1:A10` is the range of cells containing the observed (measured) values and `B1::B10` the estimated (predicted) values. Other GoF functions are used in the same way, except for AIC and BIC functions.
* To use the AIC function, type in `=fit_aic(A1:A10, B1:B10, 3, True)` where `A1:A10` and `B1:B10` contain the observed and estimated values, respectively; the third argument (value `3`) is the number of model parameters plus one (e.g., simple linear regression equation y = mx + c has 3 model parameters: m, c, and plus one); and the last parameter is True (by default) for second-order AIC. Set to False for first order AIC (use for large samples).
* The BIC function is used in the same way as the AIC function, except the BIC function is `fit_bic` and it has no fourth parameter, e.g., `=fit_bic(A1:A10, B1:B10, 3)`.
