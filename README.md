# VBA-  Excel

      VBA,            Excel.

##  

1. ** **  Excel:
   -  **** -> **** -> **  **.
   -   **   **   .
2. **  VBA**  `Alt + F11`.
3. ** **   `src`  :
   -     ,       ** ...**.
   -  `Permissions.bas`, `Logging.bas`, `Utility.bas`, `frmLogin.frm`  `ThisWorkbook.cls`.
4.        (`.xlsm`).

##   

1.   VBA         ** ...**.
2.    `LoginForm.frm`   `src` (   ,   ).
3.              .

###    

          `ThisWorkbook`.         .

## 

- **Permissions**       ,     .
- **Logging**        .
- **Utility**    :   ,  ,    .
- **frmLogin**        .

###    `ThisWorkbook`
```vba
Private Sub Workbook_Open()
    InitializeGlobals
    frmLogin.Show
    If frmLogin.IsAuthenticated Then
        Dim info As Variant
        info = frmLogin.UserInfo
        ApplyPermissions frmLogin.Username, info(1), info(2), info(3)
        LogLogin frmLogin.Username, "open", ""
    Else
        ThisWorkbook.Close False
    End If
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    LogChangeForCell Target, ""
End Sub
```

##  

- **""**   . :
  1. `A`   .
  2. `B`   (  ).
  3. `C`   (`admin`  `user`).
  4. `D`     `;`  `*`  .
  5. `E`     `Sheet!A1:B2`.
- **""**   .  :  , ,  (`ok`  `fail`), .
- **""**   .   , , ,  ,    .
- **"LockedCells"**       .       .

##  

###  
1.       ""    `admin`.
2.      `B1`  "".
3.         .

###  
1.      .
2.        .
3.       .

###  
```
 ->   ->    ""
             ->    ApplyPermissions
             ->      
  -> Logging.bas    ""
```


## 

    [Unlicense](LICENSE).

## 

    .  issue   pull request.
