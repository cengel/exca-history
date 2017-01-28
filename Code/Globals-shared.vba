Option Compare Database
Option Explicit
Public Const DBName = "Excavation Central Database"
Public VersionNumber
Public Const VersionNumberLocal = "18.2" 'NEW 2009 TO FLAG UPDATE MESSAGE TO USER - see SetCurrentVersion in module General Procedures-shared
Public GeneralPermissions
Public Const ImageLocationOnSite = "H:\Catalhoyuk\images\"
Public Const ImageLocationOnWeb = "http://www.catalhoyuk.com/database/database_new/test/getphoto.asp"
Public spString 'var to hold call to sp used in Delete_Category_SubTable_Entry() on Unit Sheet
Public Const sketchpath = "\\catal\Site_Sketches\"
Public ThisYear
