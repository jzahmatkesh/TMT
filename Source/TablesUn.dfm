object DmTables: TDmTables
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  Height = 548
  Width = 759
  object ADOConnection1: TADOConnection
    ConnectionString = 
      'Provider=SQLOLEDB.1;Password=Sanyar@jz@ss;Persist Security Info=' +
      'True;User ID=sa;Initial Catalog=TMT;Data Source=10.211.55.2'
    LoginPrompt = False
    Provider = 'SQLOLEDB.1'
    Left = 40
    Top = 24
  end
  object ADOQuery1: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    Left = 48
    Top = 80
  end
  object AdqSahamPic: TADOQuery
    Connection = ADOConnection1
    Parameters = <
      item
        Name = 'Param1'
        Attributes = [paNullable, paLong]
        DataType = ftVarBytes
        NumericScale = 255
        Precision = 255
        Size = -1
        Value = Null
      end
      item
        Name = 'Param2'
        Attributes = [paSigned, paNullable]
        DataType = ftInteger
        Precision = 10
        Size = 4
        Value = Null
      end
      item
        Name = 'Param3'
        Attributes = [paSigned, paNullable]
        DataType = ftSmallint
        Precision = 5
        Size = 2
        Value = Null
      end>
    SQL.Strings = (
      'Update TBSaham'
      'Set Pic = :Pic'
      'Where ID = :ID And Radif = :Radif')
    Left = 48
    Top = 144
  end
  object ADOExcelConnection: TADOConnection
    LoginPrompt = False
    Left = 152
    Top = 24
  end
  object AdqExcel: TADOQuery
    Connection = ADOExcelConnection
    CursorType = ctStatic
    Parameters = <>
    Left = 152
    Top = 76
  end
  object DsExcel: TDataSource
    DataSet = AdqExcel
    Left = 152
    Top = 124
  end
  object AdqBlockPic: TADOQuery
    Connection = ADOConnection1
    Parameters = <
      item
        Name = 'Pic'
        Attributes = [paNullable, paLong]
        DataType = ftVarBytes
        NumericScale = 255
        Precision = 255
        Size = 2147483647
        Value = Null
      end
      item
        Name = 'ID'
        Attributes = [paSigned]
        DataType = ftSmallint
        Precision = 5
        Size = 2
        Value = Null
      end>
    SQL.Strings = (
      'Update TBBlock'
      'Set Pic = :Pic'
      'Where ID = :ID')
    Left = 48
    Top = 208
  end
end
