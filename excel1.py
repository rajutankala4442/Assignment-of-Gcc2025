from operator import index
import xlsxwriter
 
data = [
    {
        'name': "Nagaraju",
        'roll number': "216Q1A4442",
        'phone': "8790565801",
        
    },
    
    {
        'name': "Sai Teja Alle",
        'roll number': "21B21A4594",
        'phone': "7993159368",
        'email': "allesaiteja9542@gmail.com",
        'village': "Kunthalagudem",
        'district': "Eluru"
    },
    {
        'name': "Hanuman Sai",
        'roll number': "21B21A4581",
        'phone': "9346270462",
        'email': "Khanumansaikishore@gmail.com",
        'village': "Jaganadhapuram",
        'district': "Rajhamundry"
    },
    {
        'name': "Karun Kumar",
        'roll number': "21B21A4595",
        'phone': "7569509368",
        'email': "karunkumarankem@gmail.com",
        'village': "Yedhavolu",
        'district': "Eluru"
    },
    {
         'name': "Bhaskar Anand",
        'roll number': "21B21A4588",
        'phone': "9948224752",
        'email': "bhaskaranand@gmail.com",
        'village': "Jangareddygudem",
        'district': "Eluru"
    }
]

workbook = xlsxwriter.Workbook("AllAboutPythonExcel.xlsx")
worksheet= workbook.add_worksheet("firstsheet")

worksheet.write(0, 0, "#")
worksheet.write(0, 1, "Name")
worksheet.write(0, 2, "Roll number")
worksheet.write(0, 3, "Phone")
worksheet.write(0, 4, "Email")
worksheet.write(0, 5, "Village")
worksheet.write(0, 6, "District")

for index, entry in enumerate(data):
    worksheet.write(index+1, 0, str(data))
    worksheet.write(index+1, 1, entry["name"])
    worksheet.write(index+1, 2, entry["roll number"])
    worksheet.write(index+1, 3, entry["phone"])
    worksheet.write(index+1, 4, entry["email"])
    worksheet.write(index+1, 5, entry["village"])
    worksheet.write(index+1, 6, entry["district"])

workbook.close()
