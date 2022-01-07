import pandas as pd
import sys
import mysql.connector as mysql
from tkinter import * 
from tkinter import messagebox



class Convert_ExcelSheet_To_MySqlTable:
       def __init__(self,username,password,hostname,excel_file,database_name,number_of_columns,sql):
           self.excel_file=excel_file
           self.username=username
           self.password=password
           self.hostname=hostname
           self.sql=sql
           self.database_name=database_name
           self.number_of_columns=number_of_columns
           
       def Convert_to_MySqlDatabaseTable(self):
            root = Tk()
            root.withdraw()
            read=pd.read_excel(self.excel_file, engine='openpyxl')
            read_array=read.to_numpy()
            try:
                con=mysql.connect(user=self.username,password=self.password,host=self.hostname)
                messagebox.showinfo("Authentification", "You have Logged into the database successfully")
                for reading in read_array:
                      finalvalues=[]
                      count=0
                      while count<self.number_of_columns:
                           finalvalues.append(str(reading[count]))
                           count+=1
                      cursor=con.cursor()
                      cursor.execute("USE {}".format(self.database_name))
                      #(serialnumber,entrynumber,volumenumber,district,year,user,hospital)
                      #sql=" INSERT INTO db4(serialnumber,entrynumber,volumenumber,district,year,user,hospital) VALUES(%s,%s,%s,%s,%s,%s,%s)"
                      sql=self.sql
                      cursor.execute(sql,finalvalues)
                      con.commit()
                messagebox.showinfo("Completed Inserting Data to Mysql Database","All the data in the Excelsheet has been entered into the database successfully!!!")
   
            except mysql.Error as e:
                      messagebox.showinfo("Authentifiacation","Authentifiacation Error:"+ str(e))
                      print("Authentifiacation Error:"+ str(e))
                      print("######################################################")
                      print("Make sure you enter the full path to your excel file .")
                      print("########################################################")
            root.mainloop() 



#Enter your mysql database table fields in a tuple in single quotes only '' :Note you must use only single quotes in a tuple  of your fields.
#excel_file_path=open(r"C:\Users\LAMECK\Desktop\Excel Converter\documentation\db4.xlsx","rb")

#Enter your table name and the columns you have in your mysql table .Note the columns in the excel sheet and mysql table should be in the same order.
#Enter( %s ) according to the number of columns you have .

#sql=" INSERT INTO db4(serialnumber,entrynumber,volumenumber,district,year,user,hospital) VALUES(%s,%s,%s,%s,%s,%s,%s)"


#convert=Convert_ExcelSheet_To_MySqlTable(username="root",password="tangimeko7583",hostname="localhost",excel_file=excel_file_path,database_name="db4",number_of_columns=7,sql=sql)
#convert.Convert_to_MySqlDatabaseTable()



