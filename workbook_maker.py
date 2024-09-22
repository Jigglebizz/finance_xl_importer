import math
import openpyxl
from openpyxl.worksheet.table import Table, TableStyleInfo
from typing import List
import copy
import pdb
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
  
  def __sub__( self, num : int ):
    i_rep = self.NameToInt( self.name )
    i_rep -= num
    return ExcelColumn( self.IntToName( i_rep ) )

  def __repr__( self ) -> str:
    return self.name
  
#------------------------------------------------------------------------------------------------
class ExcelCell:
  col : ExcelColumn
  row : int

  def __init__( self, col : ExcelColumn, row : int ) -> None:
    self.col = col
    self.row = row

  def inc_col( self ):
    ret = ExcelCell( copy.deepcopy( self.col ), self.row )
    self.col += 1
    return ret

  def __repr__( self ) -> str:
    return f'{str(self.col)}{self.row}'

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
    if worksheet_name in self.workbook.sheetnames:
      self.workbook.remove( self.workbook[ worksheet_name ] )

    sheet = self.workbook.create_sheet( worksheet_name )
    row = 1
    data_ranges = []

    simpl_name = worksheet_name.replace(' ', '').replace('-','_')
    table_end : ExcelCell = self.MakeTransactionTable( f'{simpl_name}_tx', ExcelCell( ExcelColumn( 'A' ), 1 ), sheet, new_statements )

      # # now create summary info
      # summary_row = statement_start
      # sheet[ f'H{summary_row}' ] = 'Summary'
      # sheet[ f'M{summary_row}' ] = 'By Category'
      # summary_row += 1
      # sheet[ f'H{summary_row}' ] = 'Expenses'
      # sheet[ f'I{summary_row}' ] = 'Income'
      # sheet[ f'J{summary_row}' ] = 'Total'
      # sheet[ f'K{summary_row}' ] = 'Cash Flow'

      # col = ExcelColumn( 'M' )
      # for cat in categorizer.categories.keys():
      #   sheet[ f'{str(col)}{summary_row}'] = cat
      #   col += 1
      # sheet[ f'{str(col)}{summary_row}' ] = 'Unknown'

      # summary_row += 1
      # sheet[ f'H{summary_row}' ] = f'=-SUMIFS($B{data_start}:$B{data_end},$B{data_start}:$B{data_end},"<0",$F{data_start}:$F{data_end},"*FALSE*")'
      # sheet[ f'I{summary_row}' ] = f'=SUMIFS($B{data_start}:$B{data_end},$B{data_start}:$B{data_end},">0",$F{data_start}:$F{data_end},"*FALSE*")'
      # sheet[ f'J{summary_row}' ] = f'=SUMIFS($B{data_start}:$B{data_end},$F{data_start}:$F{data_end},"*FALSE*")'
      # sheet[ f'K{summary_row}' ] = f'=SUMIFS($B{data_start}:$B{data_end},$E{data_start}:$E{data_end},"*TRUE*",$F{data_start}:$F{data_end},"*FALSE*")'

      # col = ExcelColumn( 'L' )
      # for cat in categorizer.categories.keys():
      #   sheet[ f'{str(col)}{summary_row}' ] = f'=ABS(SUMIFS($B{data_start}:$B{data_end},$D{data_start}:$D{data_end},"*{cat}*"))'
      #   col += 1
      # sheet[ f'{str(col)}{summary_row}' ] = f'=ABS(SUMIFS($B{data_start}:$B{data_end},$D{data_start}:$D{data_end},"*Unknown*"))'

    # and total summary
    # summary_start = row
    # sheet[ f'A{row}' ] = 'Category'
    # sheet[ f'B{row}' ] = 'Amount'
    # row += 1
    # sheet[ f'A{row}' ] = 'Expenses'
    # sheet[ f'B{row}' ] = '=-(' + ' + '.join([ f'SUMIFS(B{data_start}:B{data_end},B{data_start}:B{data_end},"<0",F{data_start}:F{data_end},"*FALSE*")' for data_start, data_end in data_ranges ]) + ')'
    # row += 1
    # sheet[ f'A{row}' ] = 'Income'
    # sheet[ f'B{row}' ] = '=' + ' + '.join([ f'SUMIFS(B{data_start}:B{data_end},B{data_start}:B{data_end},">0",F{data_start}:F{data_end},"*FALSE*")' for data_start, data_end in data_ranges ])
    # row += 1
    # sheet[ f'A{row}' ] = 'Total'
    # sheet[ f'B{row}' ] = '=' + ' + '.join([ f'SUMIFS(B{data_start}:B{data_end},F{data_start}:F{data_end},"*FALSE*")' for data_start, data_end in data_ranges ])
    # row += 1

    # for cat in categorizer.categories.keys():
    #   sheet[ f'A{row}' ] = cat
    #   sheet[ f'B{row}' ] = '=ABS(' + ' + '.join( [ f'SUMIFS($B{data_start}:$B{data_end},$D{data_start}:$D{data_end},"*{cat}*")' for data_start, data_end in data_ranges ] ) + ')'
    #   row += 1
    
    # sheet[ f'A{row}' ] = 'Unknown'
    # sheet[ f'B{row}' ] = '=ABS(' + ' + '.join( [ f'SUMIFS($B{data_start}:$B{data_end},$D{data_start}:$D{data_end},"*UNKNOWN*")' for data_start, data_end in data_ranges ] ) + ')'
    
    # summary_end = row
    # summary_table  = Table( displayName='Summary', ref=f'A{summary_start}:B{summary_end}' )
    # table_style = TableStyleInfo( name='TableStyleMedium9', showRowStripes=True )
    # summary_table.tableStyleInfo = table_style
    # sheet.add_table( summary_table )

  # returns the bottom right cell
  def MakeTransactionTable( self, name: str, start_cell: ExcelCell, sheet : openpyxl.worksheet.worksheet, new_statements : List[ Statement ] ) -> ExcelCell:
    all_tx = [ tx for stmt in new_statements for tx in stmt.transactions  ]

    cell = copy.deepcopy( start_cell )
    table_start = start_cell
    sheet[ str( cell.inc_col() ) ] = 'Date'
    sheet[ str( cell.inc_col() ) ] = 'Amount'
    sheet[ str( cell.inc_col() ) ] = 'Account'
    sheet[ str( cell.inc_col() ) ] = 'Description'
    sheet[ str( cell.inc_col() ) ] = 'Category'
    sheet[ str( cell.inc_col() ) ] = 'Matched Transfer'
    cell.row += 1 
    for tx in all_tx:
      cell.col = copy.deepcopy( start_cell.col )
      sheet[ str( cell.inc_col() ) ] = tx.date.strftime( '%m/%d/%Y' )

      sheet[ str( cell ) ] = float(tx.amount.AsExcel())
      sheet[ str( cell.inc_col() ) ].number_format = "$0.00"
      sheet[ str( cell.inc_col() ) ] = str( tx.account )
      sheet[ str( cell.inc_col() ) ] = tx.name
      sheet[ str( cell.inc_col() ) ] = tx.category
      sheet[ str( cell.inc_col() ) ] = 'TRUE' if tx.matched_transfer else 'FALSE'
      cell.row += 1

    cell.col -= 1
    cell.row -= 1
    table_end = cell

    pdb.set_trace()
    stmt_table  = Table( displayName=name, ref=f'{str( table_start ) }:{ str( table_end ) }' )
    table_style = TableStyleInfo( name='TableStyleMedium9', showRowStripes=True )
    stmt_table.tableStyleInfo = table_style
    sheet.add_table( stmt_table )