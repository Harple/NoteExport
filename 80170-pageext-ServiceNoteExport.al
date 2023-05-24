pageextension 80170 Noteexport extends "Service Orders"
{
    actions
    {
        addfirst(navigation)
        {
            action(ExportLinksToExcel)
            {
                Caption = 'Export Links to Excel';
                ApplicationArea = All;
                Promoted = true;
                PromotedCategory = Process;
                PromotedIsBig = true;
                Image = Export;
                trigger OnAction()
                begin
                    ExportLinksToExcel(Rec);
                end;
            }
            action(ExportNotesToExcel)
            {
                Caption = 'Export Notes to Excel';
                ApplicationArea = All;
                Promoted = true;
                PromotedCategory = Process;
                PromotedIsBig = true;
                Image = Export;
                trigger OnAction()
                begin
                    ExportNotesToExcel(Rec);
                end;
            }
        }
    }
    local procedure ExportLinksToExcel(var Item: Record "Service Header")
    var
        TempExcelBuffer: Record "Excel Buffer" temporary;
        LinksLbl: Label 'Links';
        ExcelFileName: Label 'Links_%1_%2';
        RecordLink: Record "Record Link";
    begin
        TempExcelBuffer.Reset();
        TempExcelBuffer.DeleteAll();
        TempExcelBuffer.NewRow();
        TempExcelBuffer.AddColumn(Item.FieldCaption("Bill-to Name"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Item.FieldCaption("No."), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(RecordLink.FieldCaption(URL1), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(RecordLink.FieldCaption(Description), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(RecordLink.FieldCaption(Created), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(RecordLink.FieldCaption("User ID"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        if Item.FindSet() then
            repeat
                RecordLink.Reset();
                RecordLink.SetRange("Record ID", Item.RecordId);
                RecordLink.SetRange(Type, RecordLink.Type::Link);
                if RecordLink.FindSet() then
                    repeat
                        TempExcelBuffer.NewRow();
                        TempExcelBuffer.AddColumn(Item."Bill-to Name", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                        TempExcelBuffer.AddColumn(Item."No.", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                        TempExcelBuffer.AddColumn(RecordLink.URL1, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                        TempExcelBuffer.AddColumn(RecordLink.Description, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                        TempExcelBuffer.AddColumn(RecordLink.Created, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                        TempExcelBuffer.AddColumn(RecordLink."User ID", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                    until RecordLink.Next() = 0;
            until Item.Next() = 0;
        TempExcelBuffer.CreateNewBook(LinksLbl);
        TempExcelBuffer.WriteSheet(LinksLbl, CompanyName, UserId);
        TempExcelBuffer.CloseBook();
        TempExcelBuffer.SetFriendlyFilename(StrSubstNo(ExcelFileName, CurrentDateTime, UserId));
        TempExcelBuffer.OpenExcel();
    end;

    local procedure ExportNotesToExcel(var Item: Record "Service Header")
    var
        TempExcelBuffer: Record "Excel Buffer" temporary;
        NotesLbl: Label 'Notes';
        ExcelFileName: Label 'Notes_%1_%2';
        RecordLink: Record "Record Link";
        RecordLinkMgt: Codeunit "Record Link Management";
        Blobline: text;
        bloblenth: integer;
        blobline1: Text;
        blobline2: Text;
        blobline3: Text;
        blobline4: Text;
        blobline5: Text;
        i: Integer;
        SplitTo: Integer;
        DivValue: Integer;
        ModValue: Integer;

    begin
        TempExcelBuffer.Reset();
        TempExcelBuffer.DeleteAll();
        TempExcelBuffer.NewRow();
        TempExcelBuffer.AddColumn(Item.FieldCaption("No."), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Item.FieldCaption(Description), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        //TempExcelBuffer.AddColumn(RecordLink.FieldCaption(Note), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(RecordLink.FieldCaption(Created), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(RecordLink.FieldCaption("User ID"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.addColumn(bloblenth, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(DivValue, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(ModValue, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(blobline1, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(blobline2, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(blobline3, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(blobline4, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(blobline5, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);

        if Item.FindSet() then
            repeat
                RecordLink.Reset();
                RecordLink.SetRange("Record ID", Item.RecordId);
                RecordLink.SetRange(Type, RecordLink.Type::Note);
                if RecordLink.FindSet() then
                    repeat
                        blobline := '';
                        blobline1 := '';
                        blobline2 := '';
                        blobline3 := '';
                        blobline4 := '';
                        blobline5 := '';
                        DivValue := 0;
                        TempExcelBuffer.NewRow();
                        TempExcelBuffer.AddColumn(Item."No.", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                        TempExcelBuffer.AddColumn(Item.Description, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                        TempExcelBuffer.AddColumn(RecordLink.Created, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                        TempExcelBuffer.AddColumn(RecordLink."User ID", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                        RecordLink.CalcFields(Note);
                        Blobline := RecordLinkMgt.ReadNote(RecordLink);
                        bloblenth := StrLen(blobline);
                        DivValue := StrLen(blobline) div 250;
                        ModValue := strLen(blobline) mod 250;
                        // if ModValue = 0 then
                        //     SplitTo := DivValue
                        // else
                        //     SplitTo := DivValue + 1;
                        // for i := 1 to SplitTo do begin
                        //     blobline2 := copyStr(blobline1, 1, 250);
                        //     if StrLen(blobline1) >= 251 then
                        //         blobline1 := CopyStr(blobline1, 251, StrLen(blobline1));
                        //     TempExcelBuffer.AddColumn(blobline2, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                        // end;





                        if DivValue >= 5 then begin
                            blobline1 := copystr(blobline, 1, 250);
                            TempExcelBuffer.AddColumn(Blobline1, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                            blobline2 := CopyStr(blobline, 251, 250);
                            TempExcelBuffer.AddColumn(Blobline2, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                            blobline3 := copystr(blobline, 501, 250);
                            TempExcelBuffer.AddColumn(Blobline3, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                            blobline4 := copystr(blobline, 751, 250);
                            TempExcelBuffer.AddColumn(Blobline4, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                            blobline5 := copystr(blobline, 1001, 250);
                            TempExcelBuffer.AddColumn(Blobline5, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                        end
                        else



                            if DivValue = 3 then begin
                                blobline1 := copystr(blobline, 1, 250);
                                TempExcelBuffer.AddColumn(Blobline1, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                                blobline2 := CopyStr(blobline, 251, 250);
                                TempExcelBuffer.AddColumn(Blobline2, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                                blobline3 := copystr(blobline, 501, 250);
                                TempExcelBuffer.AddColumn(Blobline3, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                                blobline4 := copystr(blobline, 751, 250);
                                TempExcelBuffer.AddColumn(Blobline4, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                            end
                            else

                                if DivValue = 2 then begin
                                    blobline1 := copystr(blobline, 1, 250);
                                    TempExcelBuffer.AddColumn(Blobline1, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                                    blobline2 := CopyStr(blobline, 251, 250);
                                    TempExcelBuffer.AddColumn(Blobline2, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                                    blobline3 := copystr(blobline, 501, 250);
                                    TempExcelBuffer.AddColumn(blobline3, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                                end
                                else
                                    if DivValue = 1 then begin
                                        blobline1 := copystr(blobline, 1, 250);
                                        TempExcelBuffer.AddColumn(Blobline1, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                                        blobline2 := CopyStr(blobline, 251, 250);
                                        TempExcelBuffer.AddColumn(Blobline2, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                                    end
                                    else
                                        if DivValue = 0 then begin
                                            blobline1 := copystr(blobline, 1, 250);
                                            TempExcelBuffer.AddColumn(Blobline1, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                                        end;


                    // TempExcelBuffer.addColumn(bloblenth, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                    // TempExcelBuffer.AddColumn(DivValue, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                    // TempExcelBuffer.AddColumn(ModValue, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                    until RecordLink.Next() = 0;
            until Item.Next() = 0;
        TempExcelBuffer.CreateNewBook(NotesLbl);
        TempExcelBuffer.WriteSheet(NotesLbl, CompanyName, UserId);
        TempExcelBuffer.CloseBook();
        TempExcelBuffer.SetFriendlyFilename(StrSubstNo(ExcelFileName, CurrentDateTime, UserId));
        TempExcelBuffer.OpenExcel();
    end;
}