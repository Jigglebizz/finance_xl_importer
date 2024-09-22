import argparse
import os
import yaml
import sys
from colorama import Fore, Style
import colorama
import pdb
from typing import Optional, List, Dict, Pattern
from datetime import datetime
from statement import Statement, Categorizer
from bank      import Bank, BankInfo
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
def ReadBankInfo() -> Dict[Bank, BankInfo]:
  banks = {}

  with open( g_WorkingDir + '\\' + g_BankConfigFile, 'r' ) as bank_file:
    bank_yaml = yaml.safe_load( bank_file )

    for bank in Bank:
      new_bank = BankInfo( bank )
      new_bank.ReadFromYaml( bank_yaml )
      banks[ bank ] = new_bank

  return banks

#------------------------------------------------------------------------------------------------
def IdentifyBank( csv_path : str, banks : Dict[ Bank, BankInfo ] ) -> Optional[Bank]:
  with open( csv_path, 'r' ) as csv:
    header = csv.readline()
    for bank, bank_info in banks.items():
      if ','.join(bank_info.headers) == header.strip():
        return bank
  return None
  

#------------------------------------------------------------------------------------------------
def ReadStatements( path_to_statements : str, banks : Dict[ Bank, BankInfo ], categorizer : Categorizer ) -> List[ Statement ]:
  statements : List[ Statement ] = []
  statement_file_list = os.listdir( path_to_statements )
  for file in statement_file_list:
    statement_full_path = f'{path_to_statements}\\{file}'
    bank = IdentifyBank( statement_full_path, banks )
    if not bank is None:
      bank_info = banks[ bank ]
      statement = Statement( bank )
      statement.Read( statement_full_path, bank_info, categorizer )
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

  banks          = ReadBankInfo()
  categorizer    = Categorizer( f'{g_WorkingDir}\\{g_CategoryConfigFile}' )
  new_statements = ReadStatements( args.statements, banks, categorizer )

  book_maker = WorkbookMaker( f'{g_WorkingDir}\\{g_WorkbookFile}' )
  book_maker.AppendStatements( new_statements, categorizer )
  book_maker.Save()