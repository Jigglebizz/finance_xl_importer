import argparse
import os
import sys
from colorama import Fore, Style
import colorama
import pdb
from typing import Optional, List, Dict, Pattern
from datetime import datetime
from statement import Statement, Categorizer
from bank      import Bank, BankInfo, BankManager
from currency  import USD
from workbook_maker import WorkbookMaker


g_WorkingDir         = '\\\\lore\home\\finances'
g_WorkbookFile       = 'Finances.xlsx'
g_BankConfigFile     = 'bank_config.yaml'
g_CategoryConfigFile = 'category_config.yaml'

sys.path.append( g_WorkingDir )

class NotAValidDirError( Exception ):
  pass

#------------------------------------------------------------------------------------------------
def type_dir_path( path : str ) -> str:
  full_dir = g_WorkingDir + '\\' + path
  if os.path.isdir( full_dir ):
    return full_dir
  else:
    raise NotAValidDirError(full_dir)


#------------------------------------------------------------------------------------------------
def ReadStatements( path_to_statements : str, bank_manager : BankManager, categorizer : Categorizer ) -> List[ Statement ]:
  statements : List[ Statement ] = []
  statement_file_list = os.listdir( path_to_statements )
  for file in statement_file_list:
    statement_full_path = f'{path_to_statements}\\{file}'
    bank_info = bank_manager.IdentifyStatement( statement_full_path )
    if bank_info is not None:
      statement = Statement( bank_info )
      statement.Read( statement_full_path, categorizer )
      statements.append( statement )

  for statement in statements:
    statement.ResolveTransfers( statements )

  return statements

#------------------------------------------------------------------------------------------------
if __name__=='__main__':
  def ParseArgs():
    parser = argparse.ArgumentParser()
    parser.add_argument( '--statements', type=type_dir_path, required="True", help='Name of directory where statements are stored')
    return parser.parse_args()
  args = ParseArgs()

  colorama.init()

  bank_manager   = BankManager( g_WorkingDir + '\\' + g_BankConfigFile )
  categorizer    = Categorizer( f'{g_WorkingDir}\\{g_CategoryConfigFile}' )
  new_statements = ReadStatements( args.statements, bank_manager, categorizer )

  book_maker = WorkbookMaker( f'{g_WorkingDir}\\{g_WorkbookFile}' )
  book_maker.AppendStatements( new_statements, categorizer )
  book_maker.Save()