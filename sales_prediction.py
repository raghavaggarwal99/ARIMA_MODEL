import pandas as pd
import numpy as np
from matplotlib import pyplot
from pandas import DataFrame
from pandas.plotting import autocorrelation_plot
from statsmodels.tsa.arima_model import ARIMA
from sklearn.metrics import mean_absolute_error
import xlwt 
from xlwt import Workbook 

# df=pd.read_excel("Sales1.ods")
df=pd.read_excel("Sales1.xlsx")

major_category=['M_RAINWEAR_IN','M_BOXER', 'M_BRIEF', 'M_H_BERMUDA','M_H_JAMAICAN','M_H_PYJAMA','M_SANDO','M_SHRUG_FS','M_T_BERMUDA','M_TEES_HS','M_TIGHTS_FS','M_TIGHTS_HS','M_TIGHTS_SL','M_T_JAMAICAN','M_T_PYJAMA','M_T_SHIRT_FS','M_VEST','M_T_SHIRT_PN_HS','M_T_SHIRT_SL']

category='M_T_SHIRT_SL'
df=df.loc[df['Store_Code'] =='V2-03']
df=df.loc[df['MAJ_CAT'] ==category]
df=df.drop(columns=['Store_Code', 'MAJ_CAT', 'REG', 'ST_NM','SEG','DIV','SUB_DIV','RNG_SEG', 'MVGR_MATRIX', '10_DIGIT', 'GEN_ART_DESC' ,'GEN_ART_STAT',	'SSN'])


# dropping the gmv and salev
for i in range(2,24):
    df=df.drop(columns=[('CM-'+str(i)+'_GM_V')])
for i in range(2,14):
    df=df.drop(columns=[('CM-'+str(i)+'_Sale_V')])
for i in range(14,24):
    df=df.drop(columns=[('CM-'+str(i)+'_SALE_V')])

df=df[df.columns[::-1]]

df=df.sum()

X = df.values
size = int(len(X))


wb = Workbook() 
  
# add_sheet is used to create sheet. 
sheet1 = wb.add_sheet(category) 
  
sheet1.write(0, 0, 'Predictions') 
sheet1.write(0, 1, 'Original') 
sheet1.write(0, 2, 'Error') 
prediction_row=1
original_row=1
error_row=1  
total_error=0
test = X[0:size]
history = [x for x in test]
predictions = list()
errors=list()
for t in range(len(test)):
       model = ARIMA(history, order=(4,0,0))
       model_fit = model.fit(disp=0)
       output = model_fit.forecast()
       yhat = output[0]
       obs = test[t]
       history.append(obs)
       if obs==0:
               yhat=0.0
       predictions.append(yhat)                   
       print('predicted=%f, expected=%f' % (yhat, obs))
       if obs!=0:
               q=abs(yhat-obs)
               q=q/obs
       else:
               q=0.0        
       sheet1.write(prediction_row,0,float(yhat)) 
       sheet1.write(original_row,1,float(obs)) 
       sheet1.write(error_row,2,float(q)) 
       prediction_row+=1
       error_row+=1
       original_row+=1
       total_error+=q
#        errors.append(q)
# print(error)
wb.save(str(category)+'.xls') 
total_error=total_error/len(test)
print('The total error is =%f' %(total_error))
pyplot.plot(test)
pyplot.plot(predictions, color='red')
pyplot.show()
