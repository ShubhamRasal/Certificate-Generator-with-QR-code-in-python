from PIL import Image
from PIL import ImageFont
from PIL import ImageDraw
import xlrd 
import qrcode
import time 


def main():
    print("----- ICCS's Certificate Creator -----")

    #Accept EXCEL file
    #C:\Users\ADMIN\Desktop\Indira_II\my_excel.xlsx
    path=input("Enter Path of Excel File\n")
    #print("Given path is "+path)

    #output direcotry path
    output_dir_path=input("Enter Output Directory Folder Path(add \ at the end)\n")

    #Accept Date From user
    date=input("Enter Date:")

    #open Excel File
    # To open Workbook 
    wb = xlrd.open_workbook(path) 
    sheet = wb.sheet_by_index(0)
    #ROWS and Columns
    R=sheet.nrows
    C=sheet.ncols

    basewidth = 100
    hsize=100;
    selectFont = ImageFont.truetype(r'C:\Users\ADMIN\Desktop\Indira_certification\Holy Union.ttf', 30)
    for i in range(1,R): 
       # print('|')
        comp=sheet.cell_value(i,0)
        name=sheet.cell_value(i,1)
        #take a image
        img = Image.open("template.jpg")
        draw = ImageDraw.Draw(img)
        draw.text( (270,280),"Participation In "+comp+" competition", (0,0,0), font=selectFont)
        draw.text( (350,410),name, (0,0,0), font=selectFont)
        draw.text( (200,500),date, (0,0,0), font=selectFont)
        
        #Make QR Code Of file Path
        QR_img = qrcode.make('https://erp.indira.com/use_name/certificate/'+name+'.pdf')
        #wpercent = (basewidth / float(QR_img.size[0]))
        #hsize = int((float(QR_img.size[1]) * float(wpercent)))
        #print(hsize)
        QR_img = QR_img.resize((basewidth, hsize), Image.ANTIALIAS)
        #img.save('resized_image.jpg')

        #paste the QR_code on Our Image
        img.paste(QR_img,(0,600))
        #save changed Image
        #img.save(r"C:\Users\ADMIN\Desktop\Certificates\\"+name+str(i)+".pdf", "PDF", resolution=100.0)
        img.save(output_dir_path+name+str(i)+".pdf", "PDF", resolution=100.0)
    
    print("Task Ended")
    input("Please Enter to Exit")

if __name__ == "__main__":
    main()
    input("Please Enter to Exit")