import xlrd
import csv

Sheet1 = "cda-export.xls"
Sheet2 = "tanda_staff_details.csv"

Book1 = xlrd.open_workbook(Sheet1)
second_sheet = []
with open(Sheet2) as f:
    Book2 = csv.reader(f)
    for row in Book2:
        second_sheet.append(row)

pay_rates = []
for i in second_sheet:
    if "Casual" in i[10]:
        pay_rates.append([ i[0], round(float(i[7])*1.25,2)],)
    else:
        pay_rates.append([ i[0], i[7]])


first_sheet = Book1.sheet_by_index(0)

del_speed = []
j=2
while j < len(first_sheet.col_values(0)):
    i=first_sheet.row_values(j)
    del_speed.append([ i[1], i[2], i[4], i[10], round((1000*float(i[4]))/(float(i[2])*float(i[10])*60),2)])
    j += 1


best_speed=0
fastest_driver=[]
for i in del_speed:
    best_speed = max( float(i[4]), best_speed )



everything = []
for i in del_speed:
    for j in pay_rates:
        if j[0] in i[0]:
            everything.append([j[0], j[1], i[2], i[4]])


for i in everything:
    if i[1] == '':
        i[1] = 30
    if i[3] == best_speed:
        fastest_driver=i


total_sav = 0

for i in everything:
    # 0 - name
    # 1 - wage
    # 2 - dist
    # 3 - speed
    t_total = round(1000*float(i[2])/float(i[3]),0)
    m_total = round(float(i[1])*t_total/3600,2)
    t_saving = t_total - round(1000*float(i[2])/float(fastest_driver[3]),0)
    m_sav1 = round(t_saving*float(i[1])/3600,2)
    m_sav2 = round(m_total - float(fastest_driver[1])*1000*float(i[2])/float(fastest_driver[3])/3600,2)
    i.append(m_sav1)
    i.append(m_sav2)
    total_sav+= m_sav2

# print(fastest_driver)
print(['Name','Wage','distance','speed','savings at speed of fastest driver','Savings if taken by fastest driver'])
for i in everything:
    print(i)

print("The fastest driver is "+fastest_driver[0]+". If every delivery was taken by this driver the store would have saved $"+str(total_sav))

with open('driver_costs.csv', 'w', newline='') as csvfile:
    writerObject = csv.writer(csvfile, delimiter=',',
                            quotechar=',', quoting=csv.QUOTE_MINIMAL)
    # writerObject.writerow(['Spam'] * 5 + ['Baked Beans'])
    # writerObject.writerow(['Spam', 'Lovely Spam', 'Wonderful Spam'])
    writerObject.writerow(['Name','Wage','distance','speed','savings at speed of fastest driver','Savings if taken by fastest driver'])
    for i in everything:
        writerObject.writerow(i)