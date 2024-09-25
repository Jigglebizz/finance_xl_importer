from enum import Enum
from typing import Optional, List, Pattern, Dict
from colorama import Fore, Style
from datetime import datetime
import yaml
import re

#------------------------------------------------------------------------------------------------
Bank = Enum(
  'Bank',
  [
    'Truist',
    'Discover',
    'PersonalLoan'
  ]
)

BankType = Enum(
  'BankType',
  [
    'Bank',
    'CreditCard'
  ]
)

#------------------------------------------------------------------------------------------------
class DateMatch:
  re_match : Pattern
  dt_fmt   : str

  def __init__( self, re_match : str, dt_fmt : str ) -> None:
    self.re_match = re.compile( re_match )
    self.dt_fmt   = dt_fmt

  def Match( self, filename : str, debug_name : str, debug_bank_name : str ) -> datetime:
    m = re.match( self.re_match, filename )
    if m is not None:
      date_str = m.group(1)
      return datetime.strptime( date_str, self.dt_fmt )
    else:
      print( f'{Fore.YELLOW}Warning: statement {debug_name} pattern for bank {debug_bank_name} did not match filename {filename}{Style.RESET_ALL}')
    return None

#------------------------------------------------------------------------------------------------
class BankInfo:
  bank               : Bank
  type               : BankType
  headers            : List[str]
  amount_idx         : int
  date_idx           : int
  date_fmt           : str
  name_idx           : int
  date_begin_pattern : Optional[DateMatch]
  date_end_pattern   : Optional[DateMatch]
  account_id_pattern : Optional[Pattern]

  def __init__( self, bank : Bank ) -> None:
    self.bank               = bank
    self.type               = BankType.Bank
    self.headers            = []
    self.amount_idx         = -1
    self.date_idx           = -1
    self.date_fmt           = '%m/%d/%Y'
    self.name_idx           = -1
    self.date_begin_pattern = None
    self.date_end_pattern   = None
    self.account_id_pattern = None

  def ReadFromYaml( self, yaml ) -> None:
    if self.bank.name in yaml.keys():
  
      bank_info_yaml = yaml[ self.bank.name ]

      if 'Headers' in bank_info_yaml.keys():
        self.headers = bank_info_yaml[ 'Headers' ]

      if 'Type' in bank_info_yaml.keys():
        for t in BankType:
          if t.name == bank_info_yaml[ 'Type' ].replace( ' ', '' ):
            self.type = t
            break

      if 'AmountHeader' in bank_info_yaml.keys():
        self.amount_idx = self.headers.index( bank_info_yaml[ 'AmountHeader' ] )

      if 'DateHeader' in bank_info_yaml.keys():
        self.date_idx = self.headers.index( bank_info_yaml[ 'DateHeader' ] )

      if 'DateFormat' in bank_info_yaml.keys():
        self.date_fmt = bank_info_yaml[ 'DateFormat' ]

      if 'NameHeader' in bank_info_yaml.keys():
        self.name_idx = self.headers.index( bank_info_yaml[ 'NameHeader' ] )
        
      if 'DateBegin' in bank_info_yaml.keys():
        date_begin = bank_info_yaml[ 'DateBegin' ]
        self.date_begin_pattern = DateMatch( date_begin[ 'Match' ], date_begin[ 'Fmt' ] )

      if 'DateEnd' in bank_info_yaml.keys():
        date_end = bank_info_yaml[ 'DateEnd' ]
        self.date_end_pattern = DateMatch( date_end[ 'Match' ], date_end[ 'Fmt' ] )

      if 'AccountId' in bank_info_yaml.keys():
        self.account_id_pattern = re.compile( bank_info_yaml[ 'AccountId' ] )


class BankManager:
  banks: Dict[ Bank, BankInfo ]

  def __init__( self, bank_config_abs_path : str ) -> None:
    self.banks = {}

    with open( bank_config_abs_path, 'r' ) as bank_file:
      bank_yaml = yaml.safe_load( bank_file )

      for bank in Bank:
        new_bank = BankInfo( bank )
        new_bank.ReadFromYaml( bank_yaml )
        self.banks[ bank ] = new_bank

  def IdentifyStatement( self, csv_path : str ) -> Optional[BankInfo]:
    with open( csv_path, 'r' ) as csv:
      header = csv.readline()
      for bank, bank_info in self.banks.items():
        if ','.join(bank_info.headers) == header.strip():
          return self.banks[ bank ]
    return None