///*
//* Copyright (c) 2022, Ilham Djumatov. All rights reserved.
//* Copyrights licensed under the GNU GPLv3.
//* See the accompanying LICENSE file for terms.
//*/

&AtClient
Var ExcelApp;
&AtClient
Var Book;
&AtClient
Var Sheet;

&AtServer
Procedure OnCreateAtServer(Cancel, StandardProcessing)
    
    // Реквизиты формы
    AttributesToAdd = New Array;
    AttributesToAdd.Добавить(New FormAttribute("FirstRow",     New TypeDescription("Number", New NumberQualifiers(10, 0)), ""));
    AttributesToAdd.Добавить(New FormAttribute("FirstCol",     New TypeDescription("Number", New NumberQualifiers(10, 0)), ""));
    AttributesToAdd.Добавить(New FormAttribute("LastRow",      New TypeDescription("Number", New NumberQualifiers(10, 0)), ""));
    AttributesToAdd.Добавить(New FormAttribute("LastCol",      New TypeDescription("Number", New NumberQualifiers(10, 0)), ""));
    AttributesToAdd.Добавить(New FormAttribute("List",         New TypeDescription("ValueList"), ""));
    ChangeAttributes(AttributesToAdd);
    
    Parameters.Свойство("FirstRow",     ThisObject["FirstRow"]);
    Parameters.Свойство("FirstCol",     ThisObject["FirstCol"]);
    Parameters.Свойство("LastRow",      ThisObject["LastRow"]);
    Parameters.Свойство("LastCol",      ThisObject["LastCol"]);
    
    If ThisObject["FirstRow"] = 0 Then
        ThisObject["FirstRow"] = 1;
    EndIf;
    
    If ThisObject["FirstCol"] = 0 Then
        ThisObject["FirstCol"] = 1;
    EndIf;
    
    ThisForm.CommandBar.Visible = Ложь;
    ThisForm.Title = "Выберите лист";
    ThisForm.AutoTitle = Ложь;
    ThisForm.WindowOpeningMode = FormWindowOpeningMode.Independent;
    
    Item = Items.Добавить("List", Type("FormTable"));
    Item.ПутьКДанным = "List";
    Item.ТолькоПросмотр = True;
    Item.КоманднаяПанель.Visible = Ложь;
    Item.SetAction("Selection", "ListSelection");
    
    Item = Items.Добавить("ListValue", Type("FormField"), Items["List"]);
    Item.ПутьКДанным = "List.Value";
    
EndProcedure

&AtClient
Procedure ListSelection(Item, SelectedValue, Field, StandardProcessing)
    
    Status("Выполняется чтение данных");
    
    FirstRow = ThisObject["FirstRow"];
	FirstCol = ThisObject["FirstCol"];
	LastRow = ThisObject["LastRow"];
	LastCol = ThisObject["LastCol"];
    
    WorksheetNumber = Item.CurrentData.Value;
    Try // Открываем лист
        Sheet = Book.WorkSheets(WorksheetNumber);
    Except
        ShowMessageBox(, "Не удалось открыть лист.");
        Return;
    EndTry;
    
    // Определение версии EXCEL.
    Version = Left(ExcelApp.Version,Найти(ExcelApp.Version,".")-1);
    ColCount = 0;
    RowCount = 0;
    If Version = "8" Then
        ColCount = Sheet.Cells.CurrentRegion.Columns.Count;
        RowCount = Sheet.Cells.CurrentRegion.Rows.Count;
    Else 
        // Метод SpecialCells не отображает только количество в области
        // если в области несколько областей, то количество получится неверным
        //ColCount = Sheet.Cells.SpecialCells(11).Column;
        //RowCount = Sheet.Cells.SpecialCells(11).Row;
        
        // Метод UsedRange количество использованных ячеек
        // если первая стрчка или колонка пропущены и то они не будут включаться в количество
        ColCount = Sheet.UsedRange.Columns.Count;
        RowCount = Sheet.UsedRange.Rows.Count;
        
        // Вычисляем правильное количество колонок и строк
        ColCount = Sheet.UsedRange.Column + Sheet.UsedRange.Columns.Count-1;
        RowCount = Sheet.UsedRange.Row + Sheet.UsedRange.Rows.Count-1;
    EndIf;
    
    If LastCol = 0 Then
        LastCol = ColCount;
    ElsIf ColCount < LastCol Then // не хватает колонок //Увеличение проверочного числа (как и захваченной области в самих документах) на дополнительные колонки
        ShowMessageBox(, "В файле не хватает колонок.");
        Return;
    EndIf;
    
    If LastRow = 0 Then
        LastRow = RowCount;
    ElsIf RowCount < LastRow Then // не хватает строк
        ShowMessageBox(, "В файле не хватает строк.");
        Return;
    EndIf;
    
    Range = Sheet.Range(Sheet.Cells(FirstRow,FirstCol), Sheet.Cells(LastRow,LastCol));
    Data = Range.Value.Unload();
    
    NotifyChoice(Data);
    
EndProcedure

&AtClient
Procedure OnOpen(Cancel)
    
    Dialog = New FileDialog(FileDialogMode.Open);
    Dialog.Title = "Выберите файл";
    Dialog.FullFileName = "";
    Dialog.Filter = "Excel документ (*.xls/*.xlsx)|*.xls?";
    Dialog.Multiselect = Ложь;
    Dialog.Directory = "С:\";
    
    Notify = New NotifyDescription("ProcessFileSelection", ThisObject);
    Try
        BeginPuttingFiles(Notify,Dialog, True);
    Except
        ErrorDescription = ErrorDescription();
        ErrorInfo = ErrorInfo();
        If Find(ErrorDescription, "32(0x00000020)") > 0 Then 
            ShowMessageBox(, "Ошибка совместного доступа к файлу. Пожалуйста сперва закройте файл.");
        Else 
            ShowMessageBox(, ErrorDescription);
        EndIf;
        Cancel = True;
    EndTry;
    
    Если Dialog.SelectedFiles.Count() = 0 Тогда 
        Cancel = True;
    КонецЕсли;
    
EndProcedure

&AtClient
Procedure ProcessFileSelection(Files, Params) Export 
    If Files <> Undefined Then 
        For Each TransferedFileDescription In Files Do 
            File = New File(TransferedFileDescription.FullName);
            Notify = New NotifyDescription("ProcessFile", ThisObject, File);
            File.НачатьПроверкуСуществования(Notify);
        EndDo;
    Иначе 
        Закрыть();
    КонецЕсли;
EndProcedure

&AtClient
Procedure ProcessFile(Exists, File) Export 
    
    If Exists Then
        
        Status("Выполняется чтение листов");
        
        FilePath = File.FullName;
        
        Try
            ExcelApp = New COMОбъект("Excel.Application");
        Except
            Закрыть();
            Return;
        EndTry;
        
        ExcelApp.DisplayAlerts = False;
        ExcelApp.FileValidation = 1;
        
        Try // Открываем файл
            Book = ExcelApp.Workbooks.Open(FilePath);
        Except
            // Debug
            ErrorInfo = ErrorInfo();
            ErrorDescription = ErrorDescription();
            // App
            ExcelApp.Quit();
            ExcelApp = NULL;
            Close();
            Return;
        EndTry;
        
        Для SheetNumber = 1 По Book.WorkSheets.Count Цикл
            ThisObject["List"].Add(SheetNumber, Book.WorkSheets(SheetNumber).Name);
        КонецЦикла;
    Else 
        ShowMessageBox(, "Файл не найден: " + File.FullName);
        Close();
    EndIf;
    
EndProcedure

&AtClient
Procedure BeforeClose(Cancel, Exit, MessageText, StandardProcessing)
    
    If Book <> Undefined Then
        Try
            Book.Close();
        Except
            // Debug
            ErrorInfo = ErrorInfo();
            ErrorDescription = ErrorDescription();
        EndTry;
    EndIf;
    Book = NULL;
    
    // App
    If ExcelApp <> Undefined Then
        Try
            ExcelApp.Quit();
        Except
            // Debug
            ErrorInfo = ErrorInfo();
            ErrorDescription = ErrorDescription();
        EndTry;
    EndIf;
    ExcelApp = NULL;
    
EndProcedure

// Прикрепим обработчики событий
#If Server Then
ThisForm.SetAction("OnCreateAtServer", "OnCreateAtServer");
ThisForm.SetAction("OnOpen", "OnOpen");
ThisForm.SetAction("BeforeClose", "BeforeClose");
#EndIf
