import pandas as pd
import xlrd
col101 =[0,1,2,3,4,5,6,7]
cse_101= pd.read_excel(r'E:\Year-1-Sem-1-2016.xls',skiprows=59,usecols=col101 ,nrows=62)

temp101 =[]
sum101=[]
for i in range(62):
    j =int( cse_101.iloc[i:i+1]['Difference'])
    if (j>12):

        a = int(abs(cse_101.iloc[i:i + 1]['Internal'] - cse_101.iloc[i:i + 1]['3rd Examiner']))
        b = int(abs(cse_101.iloc[i:i + 1]['External'] - cse_101.iloc[i:i + 1]['3rd Examiner']))

        if a<b:
            temp101.append(int((cse_101.iloc[i:i+1]['Internal'] + cse_101.iloc[i:i+1]['3rd Examiner'])/2))
        else:
            temp101.append(int((cse_101.iloc[i:i+1]['Internal'] + cse_101.iloc[i:i+1]['3rd Examiner'])/2))

    else:
         temp101.append(int((cse_101.iloc[i:i + 1]['Internal'] + cse_101.iloc[i:i + 1]['External'])/2))
    x=int(temp101[i] ) + float(cse_101[i:i + 1]["Tutorial(40)"])
    sum101.append(x)

col103=[0,1,2,3,4,5,6,7]
cse_103= pd.read_excel(r'E:\Year-1-Sem-1-2016.xls',skiprows=125,usecols=col103 ,nrows=62)

temp103 =[]
sum103=[]
for i in range(62):
    j =int( cse_103.iloc[i:i+1]['Difference'])
    if j>12:

        a = int(abs(cse_103.iloc[i:i + 1]['Internal'] - cse_103.iloc[i:i + 1]['3rd Examiner']))
        b = int(abs(cse_103.iloc[i:i + 1]['External'] - cse_103.iloc[i:i + 1]['3rd Examiner']))

        if a<b:
            temp103.append(int((cse_103.iloc[i:i+1]['Internal'] + cse_103.iloc[i:i+1]['3rd Examiner'])/2))
        else:
            temp103.append(int((cse_103.iloc[i:i+1]['Internal'] + cse_103.iloc[i:i+1]['3rd Examiner'])/2))

    else:
         temp103.append(int((cse_103.iloc[i:i+1]['Internal'] + cse_103.iloc[i:i+1]['External'])/2))
    x=int(temp103[i] )+ int(cse_103[i:i+1]['Tutorial'] )
    sum103.append(x)

col105 = [0, 1, 2, 3, 4, 5, 6, 7];
cse_105= pd.read_excel(r'E:\Year-1-Sem-1-2016.xls',skiprows=192,usecols=col105 ,nrows=62)

temp105 = []
sum105 = []
for i in range(62):
    j = int(cse_105.iloc[i:i + 1]['Difference'])
    if (j > 12):
         a = int(abs(cse_105.iloc[i:i + 1]['Internal'] - cse_105.iloc[i:i + 1]['3rd Examiner']))
         b = int(abs(cse_105.iloc[i:i + 1]['External'] - cse_105.iloc[i:i + 1]['3rd Examiner']))

         if a < b:
             temp105.append(int((cse_105.iloc[i:i + 1]['Internal'] + cse_105.iloc[i:i + 1]['3rd Examiner']) / 2))
         else:
             temp105.append(int((cse_105.iloc[i:i + 1]['Internal'] + cse_105.iloc[i:i + 1]['3rd Examiner']) / 2))

    else:
         temp105.append(int((cse_105.iloc[i:i + 1]['Internal'] + cse_105.iloc[i:i + 1]['External']) / 2))
    x = int(temp105[i] ) + float(cse_105[i:i + 1]['Tutorial'] )
    sum105.append(x)

col107 = [0, 1, 2, 3, 4, 5, 6, 7];
cse_107= pd.read_excel(r'E:\Year-1-Sem-1-2016.xls',skiprows=259,usecols=col107 ,nrows=62)

temp107 = [];
sum107 = [];
for i in range(62):
    j = int(cse_107.iloc[i:i + 1]['Difference'])
    if (j > 12):
         a = int(abs(cse_107.iloc[i:i + 1]['Internal'] - cse_107.iloc[i:i + 1]['3rd Examiner']))
         b = int(abs(cse_107.iloc[i:i + 1]['External'] - cse_107.iloc[i:i + 1]['3rd Examiner']))

         if a < b:
             temp107.append(int((cse_107.iloc[i:i + 1]['Internal'] + cse_107.iloc[i:i + 1]['3rd Examiner']) / 2))
         else:
             temp107.append(int((cse_107.iloc[i:i + 1]['Internal'] + cse_107.iloc[i:i + 1]['3rd Examiner']) / 2))

    else:
         temp107.append(int((cse_107.iloc[i:i + 1]['Internal'] + cse_107.iloc[i:i + 1]['External']) / 2))
    x = int(temp107[i]) + float(cse_107[i:i + 1]['Tutorial'])
    sum107.append(x)

col109 = [0, 1, 2, 3, 4, 5, 6, 7]
cse_109= pd.read_excel(r'E:\Year-1-Sem-1-2016.xls',skiprows=326,usecols=col109 ,nrows=62)

temp109 = []
sum109 = []
for i in range(62):
    j = int(cse_109.iloc[i:i + 1]['Difference'])
    if (j > 12):
         a = int(abs(cse_109.iloc[i:i + 1]['Internal'] - cse_109.iloc[i:i + 1]['3rd Examiner']))
         b = int(abs(cse_109.iloc[i:i + 1]['External'] - cse_109.iloc[i:i + 1]['3rd Examiner']))

         if a < b:
             temp109.append(int((cse_109.iloc[i:i + 1]['Internal'] + cse_109.iloc[i:i + 1]['3rd Examiner']) / 2))
         else:
             temp109.append(int((cse_109.iloc[i:i + 1]['Internal'] + cse_109.iloc[i:i + 1]['3rd Examiner']) / 2))

    else:
        temp109.append(int((cse_109.iloc[i:i + 1]['Internal'] + cse_109.iloc[i:i + 1]['External']) / 2))
    x = int(temp109[i]) + float(cse_109[i:i + 1]['Tutorial'])
    sum109.append(x)

col111 = [0, 1, 2, 3, 4, 5, 6, 7];
cse_111= pd.read_excel(r'E:\Year-1-Sem-1-2016.xls',skiprows=393,usecols=col111 ,nrows=62)

temp111 = [];
sum111= [];
for i in range(62):
    j = int(cse_111.iloc[i:i + 1]['Difference'])
    if (j > 12):
         a = int(abs(cse_111.iloc[i:i + 1]['Internal'] - cse_111.iloc[i:i + 1]['3rd Examiner']))
         b = int(abs(cse_111.iloc[i:i + 1]['External'] - cse_111.iloc[i:i + 1]['3rd Examiner']))

         if a < b:
             temp111.append(int((cse_111.iloc[i:i + 1]['Internal'] + cse_111.iloc[i:i + 1]['3rd Examiner']) / 2))
         else:
             temp111.append(int((cse_111.iloc[i:i + 1]['Internal'] + cse_111.iloc[i:i + 1]['3rd Examiner']) / 2))

    else:
        temp111.append(int((cse_111.iloc[i:i + 1]['Internal'] + cse_111.iloc[i:i + 1]['External']) / 2))
    x = int(temp111[i]) + float(cse_111[i:i + 1]['Tutorial'])
    sum111.append(x)

col108 = [0, 1, 2, 3, 4, 5];
cse_108= pd.read_excel(r'E:\Year-1-Sem-1-2016.xls',skiprows=461,usecols=col108 ,nrows=62)

temp108 = []
sum108= [];
for i in range(62):
    x = int(cse_108[i:i + 1]['CSE-108'])
    sum108.append(x)

col114 = [0, 1, 2, 3, 4, 5];
cse_114= pd.read_excel(r'E:\Year-1-Sem-1-2016.xls',skiprows=529,usecols=col114 ,nrows=62)

temp114 = [];
sum114= [];
for i in range(62):
    x = float(cse_114[i:i + 1]['CSE-114'])
    sum114.append(x)

col100 = [0, 1, 2, 3];
cse_100= pd.read_excel(r'E:\Year-1-Sem-1-2016.xls',skiprows=598,usecols=col100 ,nrows=62)

temp100 = [];
sum100= [];
for i in range(62):
    x =  float(cse_100[i:i + 1]['CSE-100'])
    sum100.append(x)

temp201 = []
for k in range(62):
    if sum101[k] >= 80:
        temp201.append(4.00)
    elif sum101[k] >= 75:
        temp201.append(3.75)
    elif sum101[k] >= 70:
        temp201.append(3.50)
    elif sum101[k] >= 65:
        temp201.append(3.25)
    elif sum101[k] >= 60:
        temp201.append(3.00)
    elif sum101[k] >= 55:
        temp201.append(2.75)
    elif sum101[k] >= 50:
        temp201.append(2.50)
    elif sum101[k] >= 45:
        temp201.append(2.25)
    elif sum101[k] >= 40:
        temp201.append(2.00)
    else:
        temp201.append(0.0)

temp203=[]
for k in range(62):
    if sum103[k]>=80 :
        temp203.append(4.00)
    elif sum103[k]>=75:
        temp203.append(3.75)
    elif sum103[k]>=70:
        temp203.append(3.50)
    elif sum103[k]>=65:
        temp203.append(3.25)
    elif sum103[k]>=60:
        temp203.append(3.00)
    elif sum103[k]>=55:
        temp203.append(2.75)
    elif sum103[k]>=50:
        temp203.append(2.50)
    elif sum103[k]>=45:
        temp203.append(2.25)
    elif sum103[k]>=40:
        temp203.append(2.00)
    else:
        temp203.append(0.0)

#print("course No 203 : ")
#print(temp203)


temp205 = []
for k in range(62):
    if sum105[k] >= 80:
        temp205.append(4.00)
    elif sum105[k] >= 75:
        temp205.append(3.75)
    elif sum105[k] >= 70:
        temp205.append(3.50)
    elif sum105[k] >= 65:
        temp205.append(3.25)
    elif sum105[k] >= 60:
        temp205.append(3.00)
    elif sum105[k] >= 55:
        temp205.append(2.75)
    elif sum105[k] >= 50:
        temp205.append(2.50)
    elif sum105[k] >= 45:
        temp205.append(2.25)
    elif sum105[k] >= 40:
        temp205.append(2.00)
    else:
        temp205.append(0.0)

#print("course No 205 : ")
#print(temp205)

temp207 = []
for k in range(62):
    if sum107[k] >= 80:
        temp207.append(4.00)
    elif sum107[k] >= 75:
        temp207.append(3.75)
    elif sum107[k] >= 70:
        temp207.append(3.50)
    elif sum107[k] >= 65:
        temp207.append(3.25)
    elif sum107[k] >= 60:
        temp207.append(3.00)
    elif sum107[k] >= 55:
        temp207.append(2.75)
    elif sum107[k] >= 50:
        temp207.append(2.50)
    elif sum107[k] >= 45:
        temp207.append(2.25)
    elif sum107[k] >= 40:
        temp207.append(2.00)
    else:
        temp207.append(0.0)

#print("course No 207 : ")
#print(temp207)

temp209 = []
for k in range(62):
    if sum109[k] >= 80:
        temp209.append(4.00)
    elif sum109[k] >= 75:
        temp209.append(3.75)
    elif sum109[k] >= 70:
        temp209.append(3.50)
    elif sum109[k] >= 65:
        temp209.append(3.25)
    elif sum109[k] >= 60:
        temp209.append(3.00)
    elif sum109[k] >= 55:
        temp209.append(2.75)
    elif sum109[k] >= 50:
        temp209.append(2.50)
    elif sum109[k] >= 45:
        temp209.append(2.25)
    elif sum109[k] >= 40:
        temp209.append(2.00)
    else:
        temp209.append(0.0)

#print("course No 201 : ")
#print(temp209)

temp211 = []
for k in range(62):
    if 2*sum111[k] >= 80:
        temp211.append(4.00)
    elif 2*sum111[k] >= 75:
        temp211.append(3.75)
    elif 2*sum111[k] >= 70:
        temp211.append(3.50)
    elif 2*sum111[k] >= 65:
        temp211.append(3.25)
    elif 2*sum111[k] >= 60:
        temp211.append(3.00)
    elif 2*sum111[k] >= 55:
        temp211.append(2.75)
    elif 2*sum111[k] >= 50:
        temp211.append(2.50)
    elif 2*sum111[k] >= 45:
        temp211.append(2.25)
    elif 2*sum111[k] >= 40:
        temp211.append(2.00)
    else:
        temp211.append(0.0)

temp208 = []
for k in range(62):
    if sum108[k] >= 80:
        temp208.append(4.00)
    elif sum108[k] >= 75:
        temp208.append(3.75)
    elif sum108[k] >= 70:
        temp208.append(3.50)
    elif sum108[k] >= 65:
        temp208.append(3.25)
    elif sum108[k] >= 60:
        temp208.append(3.00)
    elif sum108[k] >= 55:
        temp208.append(2.75)
    elif sum108[k] >= 50:
        temp208.append(2.50)
    elif sum108[k] >= 45:
        temp208.append(2.25)
    elif sum108[k] >= 40:
        temp208.append(2.00)
    else:
        temp208.append(0.0)


temp214 = []
for k in range(62):
    if 2*sum114[k] >= 80:
        temp214.append(4.00)
    elif 2*sum114[k] >= 75:
        temp214.append(3.75)
    elif 2*sum114[k] >= 70:
        temp214.append(3.50)
    elif 2*sum114[k] >= 65:
        temp214.append(3.25)
    elif 2*sum114[k] >= 60:
        temp214.append(3.00)
    elif 2*sum114[k] >= 55:
        temp214.append(2.75)
    elif 2*sum114[k] >= 50:
        temp214.append(2.50)
    elif 2*sum114[k] >= 45:
        temp214.append(2.25)
    elif 2*sum114[k] >= 40:
        temp214.append(2.00)
    else:
        temp214.append(0.0)

temp200 = []
for k in range(62):
    if 2*sum100[k] >= 80:
        temp200.append(4.00)
    elif 2*sum100[k] >= 75:
        temp200.append(3.75)
    elif 2*sum100[k] >= 70:
        temp200.append(3.50)
    elif 2*sum100[k] >= 65:
        temp200.append(3.25)
    elif 2*sum100[k] >= 60:
        temp200.append(3.00)
    elif 2*sum100[k] >= 55:
        temp200.append(2.75)
    elif 2*sum100[k] >= 50:
        temp200.append(2.50)
    elif 2*sum100[k] >= 45:
        temp200.append(2.25)
    elif 2*sum100[k] >= 40:
        temp200.append(2.00)
    else:
        temp200.append(0.0)


v= int(input("Enter roll :"))

sumcgpa=0
for z in range(62):
    m=int(cse_103.iloc[z,1:2])
    if(m==v):
        sumcgpa=((temp201[z]*3)+(temp203[z]*3)+(temp205[z]*3)+(temp207[z]*3)+(temp209[z]*3)+(temp211[z]*2)+(temp208[z]*1)+(temp214[z]*1)+(temp200[z]*1))/20
print('%.2f' % sumcgpa)

#for z in range(62):
 #   m=int(cse_103.iloc[z,1:2])
  #  sumcgpa = ((temp201[z] * 3) + (temp203[z] * 3) + (temp205[z] * 3) + (temp207[z] * 3) + (temp209[z] * 3) + (temp211[z] * 2) + (temp208[z] * 1) + (temp214[z] * 1) + (temp200[z] * 1)) / 20
   # print('%.2f' % sumcgpa)
