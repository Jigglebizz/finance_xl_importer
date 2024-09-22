import math
import openpyxl
from openpyxl.worksheet.table import Table, TableStyleInfo
from typing import List
from statement import Statement, Categorizer

#------------------------------------------------------------------------------------------------
class ExcelColumn:
  name : str

  def __init__( self, name :str ) -> None:
    self.name = name

  def NameToInt( self, name : str ) -> int:
    row_idx = 1
    i = 0
    for c in name:
      row_idx += ( ord(c) - 65 ) + i * 26
      i += 1
    return row_idx
  
  def IntToName( self, i : int ) -> str:
    name : str = ''
    while i != 0:
      name = chr(((i-1) % 26)+65) + name
      i = math.floor( (i - 1 ) / 26 )
    return name
  
  def __add__( self, num : int ):
    i_rep = self.NameToInt( self.name )
    i_rep += num
    return ExcelColumn( self.IntToName( i_rep ) )
  
  def __repr__( self ) -> str:
    return self.name

#------------------------------------------------------------------------------------------------
class WorkbookMaker:
  path     : str
  workbook : openpyxl.Workbook

  def __init__( self, path ) -> None:
    self.path     = path
    self.workbook = openpyxl.load_workbook( path )

  def Save( self ) -> None:
    self.workbook.save( self.path )

  def AppendStatements( self, new_statements : List[ Statement ], categorizer : Categorizer ) -> None:
    min_start_date = min( [ stmt.start_date for stmt in new_statements ] )
    max_end_date   = max( [ stmt.end_date   for stmt in new_statements ] )
    start_date_str = min_start_date.strftime( '%b%d %Y' )
    end_date_str   = max_end_date.strftime( '%b%d %Y')

    worksheet_name = f'{ start_date_str }-{ end_date_str }'
    print( worksheet_name )
    if worksheet_name in self.workbook.sheetnames:
      self.workbook.remove( self.workbook[ worksheet_name ] )

    sheet = self.workbook.create_sheet( worksheet_name )
    row = 1
    data_ranges = []
    for stmt in new_statements:
      account_name : str = stmt.bank.name + (' ' + str(stmt.account_id) if stmt.account_id != -1 else '')
      sheet[ f'A{row}' ] = account_name
      statement_start = row
      row += 1
      sheet[ f'A{row}' ] = 'Date'
      sheet[ f'B{row}' ] = 'Amount'
      sheet[ f'C{row}' ] = 'Description'
      sheet[ f'D{row}' ] = 'Category'
      sheet[ f'E{row}' ] = 'Include in Cash Flow'
      sheet[ f'F{row}' ] = 'Matched Transfer'
      data_start = row
      row += 1 
      for tx in stmt.transactions:
        sheet[ f'A{row}' ] = tx.date.strftime( '%m/%d/%Y' )
        sheet[ f'B{row}' ] = float(tx.amount.AsExcel())
        sheet[ f'B{row}' ].number_format = "$0.00"
        sheet[ f'C{row}' ] = tx.name
        sheet[ f'D{row}' ] = tx.category
        sheet[ f'E{row}' ] = 'TRUE'
        sheet[ f'F{row}' ] = 'TRUE' if tx.matched_transfer else 'FALSE'
        row += 1

      data_end = row - 1
      data_ranges.append( ( data_start, data_end ) )
      row += 1

      stmt_table  = Table( displayName=account_name.replace(' ', ''), ref=f'A{data_start}:F{data_end}' )
      table_style = TableStyleInfo( name='TableStyleMedium9', showRowStripes=True )
      stmt_table.tableStyleInfo = table_style
      sheet.add_table( stmt_table )

      # now create summary info
      summary_row = statement_start
      sheet[ f'H{summary_row}' ] = 'Summary'
      sheet[ f'M{summary_row}' ] = 'By Category'
      summary_row += 1
      sheet[ f'H{summary_row}' ] = 'Expenses'
      sheet[ f'I{summary_row}' ] = 'Income'
      sheet[ f'J{summary_row}' ] = 'Total'
      sheet[ f'K{summary_row}' ] = 'Cash Flow'

      col = ExcelColumn( 'M' )
      for cat in categorizer.categories.keys():
        sheet[ f'{str(col)}{summary_row}'] = cat
        col += 1
      sheet[ f'{str(col)}{summary_row}' ] = 'Unknown'

      summary_row += 1
      sheet[ f'H{summary_row}' ] = f'=-SUMIFS($B{data_start}:$B{data_end},$B{data_start}:$B{data_end},"<0",$F{data_start}:$F{data_end},"*FALSE*")'
      sheet[ f'I{summary_row}' ] = f'=SUMIFS($B{data_start}:$B{data_end},$B{data_start}:$B{data_end},">0",$F{data_start}:$F{data_end},"*FALSE*")'
      sheet[ f'J{summary_row}' ] = f'=SUMIFS($B{data_start}:$B{data_end},$F{data_start}:$F{data_end},"*FALSE*")'
      sheet[ f'K{summary_row}' ] = f'=SUMIFS($B{data_start}:$B{data_end},$E{data_start}:$E{data_end},"*TRUE*",$F{data_start}:$F{data_end},"*FALSE*")'

      col = ExcelColumn( 'L' )
      for cat in categorizer.categories.keys():
        sheet[ f'{str(col)}{summary_row}' ] = f'=ABS(SUMIFS($B{data_start}:$B{data_end},$D{data_start}:$D{data_end},"*{cat}*"))'
        col += 1
      sheet[ f'{str(col)}{summary_row}' ] = f'=ABS(SUMIFS($B{data_start}:$B{data_end},$D{data_start}:$D{data_end},"*Unknown*"))'

    # and total summary
    summary_start = row
    sheet[ f'A{row}' ] = 'Category'
    sheet[ f'B{row}' ] = 'Amount'
    row += 1
    sheet[ f'A{row}' ] = 'Expenses'
    sheet[ f'B{row}' ] = '=-(' + ' + '.join([ f'SUMIFS(B{data_start}:B{data_end},B{data_start}:B{data_end},"<0",F{data_start}:F{data_end},"*FALSE*")' for data_start, data_end in data_ranges ]) + ')'
    row += 1
    sheet[ f'A{row}' ] = 'Income'
    sheet[ f'B{row}' ] = '=' + ' + '.join([ f'SUMIFS(B{data_start}:B{data_end},B{data_start}:B{data_end},">0",F{data_start}:F{data_end},"*FALSE*")' for data_start, data_end in data_ranges ])
    row += 1
    sheet[ f'A{row}' ] = 'Total'
    sheet[ f'B{row}' ] = '=' + ' + '.join([ f'SUMIFS(B{data_start}:B{data_end},F{data_start}:F{data_end},"*FALSE*")' for data_start, data_end in data_ranges ])
    row += 1
    sheet[ f'A{row}' ] = 'Cash Flow'
    sheet[ f'B{row}' ] = '=' + ' + '.join([ f'SUMIFS(B{data_start}:B{data_end},E{data_start}:E{data_end},"*TRUE*")' for data_start, data_end in data_ranges ])
    row += 1

    for cat in categorizer.categories.keys():
      sheet[ f'A{row}' ] = cat
      sheet[ f'B{row}' ] = '=ABS(' + ' + '.join( [ f'SUMIFS($B{data_start}:$B{data_end},$D{data_start}:$D{data_end},"*{cat}*")' for data_start, data_end in data_ranges ] ) + ')'
      row += 1
    
    sheet[ f'A{row}' ] = 'Unknown'
    sheet[ f'B{row}' ] = '=ABS(' + ' + '.join( [ f'SUMIFS($B{data_start}:$B{data_end},$D{data_start}:$D{data_end},"*UNKNOWN*")' for data_start, data_end in data_ranges ] ) + ')'
    
    summary_end = row
    summary_table  = Table( displayName='Summary', ref=f'A{summary_start}:B{summary_end}' )
    table_style = TableStyleInfo( name='TableStyleMedium9', showRowStripes=True )
    summary_table.tableStyleInfo = table_style
    sheet.add_table( summary_table )