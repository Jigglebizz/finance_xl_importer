from enum import Enum
from typing import Optional, List, Pattern, Dict
from colorama import Fore, Style
from datetime import datetime
import re

#------------------------------------------------------------------------------------------------
Bank = Enum(
  'Bank',
  [
    'Truist',
    'Discover'
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