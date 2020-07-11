import os
import pandas as pd
import xlsxwriter
import math

def ScalingData(path,filename):
    df=pd.read_excel(path, header=None)
    # totalrows=len(df.axes[0])
    totalcols=len(df.axes[1])
    minlist=[]
    maxlist=[]

    for i in range(totalcols):
        minlist.append(df[i].min())
        maxlist.append(df[i].max())

    WorkBook=xlsxwriter.Workbook(filename)
    WorkSheet=WorkBook.add_worksheet()

    for i in range(df.shape[1]):
        collist=df.iloc[: ,i]
        row=0
        for j in collist.values:
            val=(j-minlist[i])/(maxlist[i]-minlist[i])
            WorkSheet.write(row,i,val)
            row+=1
            
    WorkBook.close()
    return df

#path of the training file
path=os.getcwd()
path=path+"\parktraining.xlsx"
#scaling tarining data
Df_train=ScalingData(path,"ScaledTrainingData.xlsx")
totalrows=(len(Df_train.axes[0]))
totalcols=(len(Df_train.axes[1]))

#path of the testing file
path=os.getcwd()
path=path+"\parktesting.xlsx"
#scaling testing data
Df_test=ScalingData(path,"ScaledTestingData.xlsx")

dict={}

classcount=Df_train.iloc[: ,totalcols-1]
index=0
for i in classcount.values:
    if i not in dict:
        dict[i]=1
        dict[str(i)+"index"]=[]
        dict[str(i)+"index"].append(index)
    else:
        dict[i]+=1
        dict[str(i)+"index"].append(index)
    index+=1


PriorProbOfClass0=(dict[0])/totalrows
PriorProbOfClass1=(dict[1])/totalrows


path=os.getcwd()
path=path+"\ScaledTrainingData.xlsx"

Df_ScaledTrain=pd.read_excel(path, header=None)

meanofClass_0=[]
meanofClass_1=[]

varOfClass_0=[]
varOfClass_1=[]

for i in range(Df_train.shape[1]):
    meanofClass_0.append(Df_ScaledTrain.iloc[dict["0index"],i].mean(axis=0))
    meanofClass_1.append(Df_ScaledTrain.iloc[dict["1index"],i].mean(axis=0))
    
    varOfClass_0.append(Df_ScaledTrain.iloc[dict["0index"],i].var(axis=0))
    varOfClass_1.append(Df_ScaledTrain.iloc[dict["1index"],i].var(axis=0))

# data trained

#working on tetsing data
path=os.getcwd()
path=path+"\ScaledTestingData.xlsx"

Df_ScaledTest=pd.read_excel(path, header=None)

WB_Ans=xlsxwriter.Workbook("Answer.xlsx")
cell_format = WB_Ans.add_format({'align': 'center'})
cell_format.set_bold()
cell_format.set_border()
WS=WB_Ans.add_worksheet('prob_0')
WSS=WB_Ans.add_worksheet('prob_1')
WS_Ans=WB_Ans.add_worksheet('answer')
WS_Ans.set_column(0, 1, 20)
WS_Ans.set_column(3, 4, 13)
WS_Ans.set_column(7, 7, 20)
WS_Ans.set_column(8, 9, 13)

WS.set_column(0,21,13)
WSS.set_column(0,21,13)

WS.set_column(23,23,20)
WSS.set_column(23,23,20)
predictedlist=[]
WS_Ans.write(0,0,"Prob(class=0/features)",cell_format)
WS_Ans.write(0,1,"Prob(class=1/features)",cell_format)
WS_Ans.write(0,3,"Predicted Class",cell_format)
WS_Ans.write(0,4,"Actual Class",cell_format)
flag=0
flagg=0
WS.write(0,23,"P(class=0/features)",cell_format)
WSS.write(0,23,"P(class=1/features)",cell_format)

for i in range((Df_ScaledTest.shape[0])):
    rowlist=Df_ScaledTest.iloc[i,Df_ScaledTest.columns != totalcols-1]
    col=0
    prob0=1
    for j in rowlist.values:
        numerator=math.exp(-( (j-meanofClass_0[col])**2 / (2*varOfClass_0[col]) ) )
        denominator=math.sqrt(2*math.pi*varOfClass_0[col])
        val=numerator/denominator
        prob0*=val
        if flag!=1:
            WS.write(0,col,"P(f"+str(col+1)+"/c=0)",cell_format)
        WS.write(i+1,col,val)
        col+=1
    
    flag=1
    WS.write(i+1,col,None,cell_format)
    col+=1
    prob0*=PriorProbOfClass0
    
    prob1=1
    col=0
    for K in rowlist.values:
        numerator=math.exp(-( (K-meanofClass_1[col])**2 / (2*varOfClass_1[col]) ) )
        denominator=math.sqrt(2*math.pi*varOfClass_1[col])
        val=numerator/denominator
        prob1*=val
        WSS.write(i+1,col,val)
        if flagg!=1:
            WSS.write(0,col,"P(f"+str(col+1)+"/c=1)",cell_format)
        col+=1
    
    flagg=1
    WSS.write(i+1,col,None,cell_format)
    col+=1
    prob1*=PriorProbOfClass1
    
    a1=(prob0)/(prob0+prob1)
    WS.write(i+1,col,a1,cell_format)
    WS_Ans.write(i+1,0,a1)
    
    a2=(prob1)/(prob0+prob1)
    WSS.write(i+1,col,a2,cell_format)
    WS_Ans.write(i+1,1,a2)
    
    if a1 > a2:
        WS_Ans.write(i+1,3,0)
        predictedlist.append(0)
    else:
        WS_Ans.write(i+1,3,1)
        predictedlist.append(1)

confusionMatrix = [[0 for i in range(len(dict)-1)] for j in range(len(dict)-1)] 
Actualresultlist=Df_ScaledTest.iloc[:,totalcols-1]

for i in range(len(Actualresultlist)):
    WS_Ans.write(i+1,4,Actualresultlist[i])
    if Actualresultlist[i]==0 and predictedlist[i]==0:
        confusionMatrix[0][0]+=1
    elif Actualresultlist[i]==0 and predictedlist[i]==1:
        confusionMatrix[0][1]+=1
    elif Actualresultlist[i]==1 and predictedlist[i]==0:
        confusionMatrix[1][0]+=1
    elif Actualresultlist[i]==1 and predictedlist[i]==1:
        confusionMatrix[1][1]+=1


confusionMatrix[0][2]=confusionMatrix[0][0]+confusionMatrix[0][1]
confusionMatrix[1][2]=confusionMatrix[1][0]+confusionMatrix[1][1]
confusionMatrix[2][0]=confusionMatrix[0][0]+confusionMatrix[1][0]
confusionMatrix[2][1]=confusionMatrix[0][1]+confusionMatrix[1][1]
confusionMatrix[2][2]=confusionMatrix[2][0]+confusionMatrix[2][1]
Accuracy=0
for i in range(len(dict)-2):
    Accuracy+=confusionMatrix[i][i]

Accuracy=Accuracy/(confusionMatrix[len(dict)-2][len(dict)-2])

# print("Total Matched :", confusionMatrix[0][0]+confusionMatrix[1][1])
# print("Total MisMatched :",confusionMatrix[0][1]+confusionMatrix[1][0])
# print ("Accuracy : ",Accuracy*100,"%")

WS_Ans.merge_range("H0:J0", "", cell_format)
WS_Ans.write(0,7,"Confusion Matrix",cell_format)

WS_Ans.write(3,7,"Actual 0",cell_format)
WS_Ans.write(4,7,"Actual 1",cell_format)
WS_Ans.write(2,8,"Predicted 0",cell_format)
WS_Ans.write(2,9,"Predicted 1",cell_format)
WS_Ans.write(3,8,confusionMatrix[0][0],cell_format)
WS_Ans.write(3,9,confusionMatrix[0][1],cell_format)
WS_Ans.write(4,8,confusionMatrix[1][0],cell_format)
WS_Ans.write(4,9,confusionMatrix[1][1],cell_format)

WS_Ans.write(5,8,confusionMatrix[2][0],cell_format)
WS_Ans.write(5,9,confusionMatrix[2][1],cell_format)
WS_Ans.write(3,10,confusionMatrix[0][2],cell_format)
WS_Ans.write(4,10,confusionMatrix[1][2],cell_format)
WS_Ans.write(5,10,confusionMatrix[2][2],cell_format)

WS_Ans.write(7,7,"Total Matched :",cell_format)
WS_Ans.write(7,8, confusionMatrix[0][0]+confusionMatrix[1][1],cell_format)

WS_Ans.write(8,7,"Total MisMatched :",cell_format)
WS_Ans.write(8,8, confusionMatrix[0][1]+confusionMatrix[1][0],cell_format)
WS_Ans.write(9,7,"Accuracy(%): ",cell_format)
WS_Ans.write(9,8,Accuracy*100,cell_format)
 
WB_Ans.close()