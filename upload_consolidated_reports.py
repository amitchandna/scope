import dropbox

def upload_to_dbox(folder, filename):
   #Lets talk about the source file
    filepath = folder / filename
    #Where we going in Drop Box
    target = "Exports"
    targetfile = target + filename
    
    dbx = dropbox.Dropbox('BzqeVllxe4AAAAAAAAAAIEeSF0ofU5aCXkNYcfknQVYmkkzBzhs-oymBELNsDy8C')
    
    with filepath.open("rb") as f:
        dbx.files_upload(f.read(),targetfile,mode=dropbox.files.WriteMode("overwrite"))

    
upload_to_dbox("some_folder", "Project_Info.xlsx")
upload_to_dbox("some_folder", "Project_Tools.xlsx")
upload_to_dbox("some_folder", "Assessments.xlsx")
upload_to_dbox("some_folder", "Assessments_Section_Response.xlsx")
upload_to_dbox("some_folder", "Account_User.xlsx")
upload_to_dbox("some_folder", "Producing_Organization_Details.xlsx")

