import pyodbc
import msaccessdb
import os,sys
if len(sys.argv)>1:
    args=sys.argv[1]
    Files1=os.listdir(args)
    file_list=[]

    for fie in Files1:
        if fie.lower().endswith(('.csv')):
                      file_list.append(fie)
                      
    for file in file_list:
        msaccessdb.create(os.path.join(args,file.split('.')[0] + ".mdb"))

        conn_str = (r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + os.path.join(args,file.split('.')[0] + ".mdb"))
        cnxn = pyodbc.connect(conn_str)
        crsr = cnxn.cursor()
        pth=args
        sql1 = "select * into " + file.split('.')[0] + " from [text;HDR=yes;Database=" +  pth + "].[" + file + "]"
        crsr.execute(sql1)
        crsr.commit()
        crsr.close()


