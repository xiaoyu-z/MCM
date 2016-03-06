__author__ = 'zhengxiaoyu'
import xlrd
workbook = xlrd.open_workbook('yourdata.xlsx')
sheet = workbook.sheets()[0]

graduate_num =  sheet.col(59+25)
public_net_price =  sheet.col(73+25)
private_net_price = sheet.col(78+25)
part_time_rate = sheet.col(69+25)
ret_rate_full_four = sheet.col(84+25)
ret_rate_full_lessfour = sheet.col(85+25)
#ret_rate_full_two = sheet.col(86)
operating = sheet.col(95)
ret_rate_part_four =sheet.col(86+25)
ret_rate_part_lessfour = sheet.col(87+25)
over25 = sheet.col(114)
loan_rate = sheet.col(113)
med_income = sheet.col(120)
comp_rate_150 = sheet.col(118)
comp_rate_200 = sheet.col(119)
stu_need_help = []
for i in range(sheet.nrows-1):
    #print ret_rate_part_lessfour[i+1],ret_rate_part_four[i+1],part_time_rate[i+1],ret_rate_full_lessfour[i+1],ret_rate_full_four[i+1]
    stu_need_help.append(int((float(part_time_rate[i+1].value)*(1-float(ret_rate_part_lessfour[i+1].value)-float(ret_rate_part_four[i+1].value))+(1-float(part_time_rate[i+1].value))*(1-float(ret_rate_full_lessfour[i+1].value)-float(ret_rate_full_four[i+1].value)))*float(graduate_num[i+1].value)*0.39))
print sum(stu_need_help)
expense = []
for i in range(len(stu_need_help)):
    expense.append(int(private_net_price[i+1].value)+int(public_net_price[i+1].value))
#28500
total_expense = []
for i in range(len(stu_need_help)):
    total_expense.append(int((stu_need_help[i]*expense[i]*4-stu_need_help[i]*(1-float(over25[i+1].value))*3800*4)*(1+1/3)- float(loan_rate[i+1].value)*stu_need_help[i]*28500))
#print total_expense
#1/3
returnfrom = []
factor_1 = 1.03 ** 9
factor_2 = 1.03 ** 5
part_time_start = []
full_time_start = []
for i in range(len(stu_need_help)):
    part_time_start.append(int(med_income[i+1].value)/factor_1)
for i in range(len(stu_need_help)):
    full_time_start.append(int(med_income[i+1].value)/factor_2)
factor_3 = (1.03**30-1)/0.03
print factor_3
#part_time 3%
#full_time 3%
#work 30 years
#all graduate
for i in range(len(stu_need_help)):
    returnfrom.append((float(part_time_rate[i+1].value)*part_time_start[i]+(1-float(part_time_rate[i+1].value))*full_time_start[i])*factor_3*0.84*stu_need_help[i]*(float(comp_rate_150[i+1].value)+float(comp_rate_200[i+1].value)))
print sum(returnfrom)

import xlwt
result = xlwt.Workbook()
table = result.add_sheet('result')
for i in range (len(stu_need_help)):
    return_value = returnfrom[i]
    expense_value = total_expense[i]
    if expense_value<0:
        expense_value = 0
    if(return_value>0):
        result_value = return_value - expense_value
    table.write(i,0,stu_need_help[i])
    table.write(i,1,expense_value)
    table.write(i,3,result_value)
    if stu_need_help[i]==0 or int(operating[i+1].value) == 0:
        table.write(i,2,0)
        table.write(i,4,0)
    else:
        expense_ave = int(expense_value/stu_need_help[i])
        if expense_ave>2000:
            table.write(i,2,expense_ave)
        else:
            table.write(i,2,0)
        table.write(i,4,int(return_value/stu_need_help[i]))
result.save('result.xls')