codeunit 50200 "Entity Text Import/Export"
{
    Subtype = Test;

    var
        Assert: Codeunit "Library Assert";
        Any: Codeunit Any;

    [Test]
    procedure TestImportEntityText()
    var
        Item: Record Item;
        EntityText: Record "Entity Text";
        TempExcelBuffer: Record "Excel Buffer" temporary;
        ItemNo: Code[20];
        ScenarioText: Text[50];
        TextContent: Text[1024];
    begin
        // Setup
        ItemNo := CreateTestItem();
        ScenarioText := 'Marketing Text';
        TextContent := 'Test marketing text for import';
        CreateTestExcelBuffer(TempExcelBuffer, ItemNo, ScenarioText, TextContent);

        // Exercise
        ImportEntityTextFromExcel(TempExcelBuffer);

        // Verify
        EntityText.SetRange("Source Table Id", Database::Item);
        EntityText.SetRange("Source System Id", GetItemSystemId(ItemNo));
        Assert.IsTrue(EntityText.FindFirst(), 'Entity Text record should be created');
        Assert.AreEqual(ScenarioText, Format(EntityText.Scenario), 'Scenario should match');
        Assert.AreEqual(TextContent, EntityText."Preview Text", 'Preview Text should match');

        // Teardown
        DeleteTestItem(ItemNo);
        DeleteEntityText(EntityText);
    end;

    [Test]
    procedure TestExportEntityText()
    var
        Item: Record Item;
        EntityText: Record "Entity Text";
        TempExcelBuffer: Record "Excel Buffer" temporary;
        ItemNo: Code[20];
        ScenarioText: Text[50];
        TextContent: Text[1024];
    begin
        // Setup
        ItemNo := CreateTestItem();
        ScenarioText := 'Marketing Text';
        TextContent := 'Test marketing text for export';
        CreateTestEntityText(ItemNo, ScenarioText, TextContent);

        // Exercise
        ExportEntityTextToExcel(TempExcelBuffer);

        // Verify
        Assert.IsTrue(TempExcelBuffer.FindSet(), 'Excel Buffer should contain exported data');
        TempExcelBuffer.Next(); // Skip header row
        Assert.AreEqual(ItemNo, GetExcelBufferCellValue(TempExcelBuffer, 2, 1), 'Exported Item No. should match');
        Assert.AreEqual(ScenarioText, GetExcelBufferCellValue(TempExcelBuffer, 2, 2), 'Exported Scenario should match');
        Assert.AreEqual(TextContent, GetExcelBufferCellValue(TempExcelBuffer, 2, 3), 'Exported Text Content should match');

        // Teardown
        DeleteTestItem(ItemNo);
    end;

    [Test]
    procedure TestClearAllEntityText()
    var
        Item: Record Item;
        EntityText: Record "Entity Text";
        ItemNo: Code[20];
    begin
        // Setup
        ItemNo := CreateTestItem();
        CreateTestEntityText(ItemNo, 'Marketing Text', 'Test marketing text');

        // Verify setup
        EntityText.SetRange("Source Table Id", Database::Item);
        EntityText.SetRange("Source System Id", GetItemSystemId(ItemNo));
        Assert.AreEqual(1, EntityText.Count, 'Should have 1 Entity Text record before clearing');

        // Exercise
        ClearAllEntityText();

        // Verify
        EntityText.Reset();
        Assert.AreEqual(0, EntityText.Count, 'Should have 0 Entity Text records after clearing');

        // Teardown
        DeleteTestItem(ItemNo);
    end;

    local procedure CreateTestItem(): Code[20]
    var
        Item: Record Item;
    begin
        Item.Init();
        Item."No." := CopyStr(Any.AlphanumericText(20), 1, 20);
        Item.Insert(true);
        exit(Item."No.");
    end;

    local procedure DeleteTestItem(ItemNo: Code[20])
    var
        Item: Record Item;
    begin
        if Item.Get(ItemNo) then
            Item.Delete(true);
    end;

    local procedure CreateTestEntityText(ItemNo: Code[20]; ScenarioText: Text[50]; TextContent: Text[1024])
    var
        EntityText: Record "Entity Text";
        Item: Record Item;
    begin
        Item.Get(ItemNo);
        EntityText.Init();
        EntityText."Source Table Id" := Database::Item;
        EntityText."Source System Id" := Item.SystemId;
        Evaluate(EntityText.Scenario, ScenarioText);
        EntityText."Preview Text" := TextContent;
        EntityText.Insert(true);
    end;

    local procedure ExportEntityTextToExcel(var TempExcelBuffer: Record "Excel Buffer" temporary)
    var
        EntityText: Record "Entity Text";
        Item: Record Item;
        RowNo: Integer;
    begin
        TempExcelBuffer.Reset();
        TempExcelBuffer.DeleteAll();

        // Add header row
        RowNo := 1;
        InsertExcelBufferCell(TempExcelBuffer, RowNo, 1, 'Item No.');
        InsertExcelBufferCell(TempExcelBuffer, RowNo, 2, 'Scenario');
        InsertExcelBufferCell(TempExcelBuffer, RowNo, 3, 'Text Content');

        // Add data rows
        EntityText.Reset();
        EntityText.SetRange("Source Table Id", Database::Item);
        if EntityText.FindSet() then
            repeat
                if Item.GetBySystemId(EntityText."Source System Id") then begin
                    RowNo += 1;
                    InsertExcelBufferCell(TempExcelBuffer, RowNo, 1, Item."No.");
                    InsertExcelBufferCell(TempExcelBuffer, RowNo, 2, Format(EntityText.Scenario));
                    InsertExcelBufferCell(TempExcelBuffer, RowNo, 3, EntityText."Preview Text");
                end;
            until EntityText.Next() = 0;
    end;

    local procedure CreateTestExcelBuffer(var TempExcelBuffer: Record "Excel Buffer" temporary; ItemNo: Code[20]; ScenarioText: Text[50]; TextContent: Text[1024])
    begin
        TempExcelBuffer.DeleteAll();
        InsertExcelBufferCell(TempExcelBuffer, 1, 1, 'Item No.');
        InsertExcelBufferCell(TempExcelBuffer, 1, 2, 'Scenario');
        InsertExcelBufferCell(TempExcelBuffer, 1, 3, 'Text Content');
        InsertExcelBufferCell(TempExcelBuffer, 2, 1, ItemNo);
        InsertExcelBufferCell(TempExcelBuffer, 2, 2, ScenarioText);
        InsertExcelBufferCell(TempExcelBuffer, 2, 3, TextContent);
    end;

    local procedure DeleteEntityText(var EntityText: Record "Entity Text")
    begin
        EntityText.DeleteAll(true);
    end;

    local procedure ImportEntityTextFromExcel(var TempExcelBuffer: Record "Excel Buffer" temporary)
    var
        EntityText: Record "Entity Text";
        Item: Record Item;
        RowNo: Integer;
        MaxRowNo: Integer;
        ScenarioValue: Enum "Entity Text Scenario";
    begin
        RowNo := 0;
        MaxRowNo := 0;
        TempExcelBuffer.Reset();
        if TempExcelBuffer.FindLast() then
            MaxRowNo := TempExcelBuffer."Row No.";

        for RowNo := 2 to MaxRowNo do begin
            Item.Get(GetExcelBufferCellValue(TempExcelBuffer, RowNo, 1));
            Evaluate(ScenarioValue, GetExcelBufferCellValue(TempExcelBuffer, RowNo, 2));

            EntityText.Reset();
            EntityText.SetRange("Source Table Id", Database::Item);
            EntityText.SetRange("Source System Id", Item.SystemId);
            EntityText.SetRange(Scenario, ScenarioValue);

            if EntityText.FindFirst() then begin
                EntityText."Preview Text" := CopyStr(GetExcelBufferCellValue(TempExcelBuffer, RowNo, 3), 1, 1024);
                EntityText.Modify();
            end else begin
                EntityText.Init();
                EntityText."Source Table Id" := Database::Item;
                EntityText."Source System Id" := Item.SystemId;
                EntityText.Scenario := ScenarioValue;
                EntityText."Preview Text" := CopyStr(GetExcelBufferCellValue(TempExcelBuffer, RowNo, 3), 1, 1024);
                EntityText.Insert();
            end;
        end;
    end;

    local procedure GetExcelBufferCellValue(var TempExcelBuffer: Record "Excel Buffer" temporary; RowNo: Integer; ColNo: Integer): Text
    begin
        TempExcelBuffer.Reset();
        TempExcelBuffer.SetRange("Row No.", RowNo);
        TempExcelBuffer.SetRange("Column No.", ColNo);
        if TempExcelBuffer.FindFirst() then
            exit(TempExcelBuffer."Cell Value as Text");
    end;

    local procedure ClearAllEntityText()
    var
        EntityText: Record "Entity Text";
    begin
        EntityText.DeleteAll(true);
    end;

    local procedure GetExcelBufferCellValue(var TempExcelBuffer: Record "Excel Buffer"; ColNo: Integer): Text
    begin
        TempExcelBuffer.SetRange("Column No.", ColNo);
        if TempExcelBuffer.FindFirst() then
            exit(TempExcelBuffer."Cell Value as Text");
    end;

    local procedure GetItemSystemId(ItemNo: Code[20]): Guid
    var
        Item: Record Item;
    begin
        Item.Get(ItemNo);
        exit(Item.SystemId);
    end;

    local procedure InsertExcelBufferCell(var TempExcelBuffer: Record "Excel Buffer" temporary; RowNo: Integer; ColNo: Integer; CellValue: Text[1024])
    begin
        TempExcelBuffer.Init();
        TempExcelBuffer.Validate("Row No.", RowNo);
        TempExcelBuffer.Validate("Column No.", ColNo);
        TempExcelBuffer.Validate("Cell Value as Text", CellValue);
        TempExcelBuffer.Insert();
    end;
}