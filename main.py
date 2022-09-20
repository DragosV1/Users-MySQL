from database import (
connect, 
create_table, 
insert_into_table,
update_name, 
update_password, 
delete_from_table, 
fetchall,
export_to_excel)


while True:
    user_input = input("\nConnect to MySQL server (C)\nInsert into table (I)\nUpdate Password (P)\nUpdate Name (N)\nDelete from table (D)\nView table (V)\nExport to Excel (E)\nQuit (Q)\nChose an option: ").upper()
    if user_input == "C":
        connect()
    elif user_input == "I":
        insert_into_table(
            name=input("Enter a name: "),
            password=input("Enter a password: ")
        )
    elif user_input == "P":
        update_password(
            id_=input("Enter the id where you want to change the password: "),
            password=input("Update password: "),\
        )
    elif user_input == "N":
        update_name(
            id_=input("Enter the id where you want to change the name: "),
            name=input("Update name: ")
        )
    elif user_input == "D":
        delete_from_table(
            id_=input("Enter the id you want to delete: ")
        )
    elif user_input == "V":
        fetchall()
    elif user_input == "E":
        try:
            export_to_excel()
            print("Exported to file users.xls")
        except:
            print("Failed to export!")
    elif user_input == "Q":
        print("\nYou have closed the program!")
        quit()
    else:
        print("\nYou entered the wrong letter!")
        quit()