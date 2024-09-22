import csv
import os
import re
import yaml
from datetime import datetime
from typing import List, Dict, Pattern
from bank   import Bank, BankInfo
from currency import USD

#------------------------------------------------------------------------------------------------
class Categorizer:
  categories : Dict[ str, List[Pattern] ]

  def __init__( self, path_to_config ) -> None:
    self.categories = {}
    with open( path_to_config, 'r' ) as config_file:
      config_yaml = yaml.safe_load( config_file )
      for category in config_yaml.keys():
        self.categories[ category ] = [ re.compile( pattern ) for pattern in config_yaml[ category ] ]

  def GetCategoryForName( self, name : str ) -> str:
    for cat, patterns in self.categories.items():
      for pattern in patterns:
        if re.search( pattern, name ) is not None:
          return cat
        
    return 'Unknown'

#------------------------------------------------------------------------------------------------
class Transaction:
  name             : str
  amount           : USD
  date             : datetime
  category         : str
  matched_transfer : bool

  def __init__( self, name : str, amount : USD, date : datetime, category : str ):
    self.name             = name
    self.amount           = amount
    self.date             = date
    self.category         = category
    self.matched_transfer = False

  def __repr__( self ) -> str:
    return f'{self.name}: {self.amount} on {self.date}'

#------------------------------------------------------------------------------------------------
class Statement:
  bank         : Bank
  start_date   : datetime
  end_date     : datetime
  account_id   : int
  transactions : List[Transaction]
  
  def __init__( self, bank : Bank ):
    self.bank         = bank
    self.transactions = []
    self.start_date   = None
    self.end_date     = None
    self.account_id   = -1

  def __repr__( self ) -> str:
    acct_str = f'acct {self.account_id} ' if self.account_id != -1 else ''
    return f'Statement for {self.bank.name} {acct_str}from {self.start_date} to {self.end_date}:\n  ' + '\n  '.join( [ str(t) for t in self.transactions] )
  
  def Read( self, statement_path : str, bank_info : BankInfo, categorizer : Categorizer ) -> None:
    file_basename = os.path.basename( statement_path )

    with open( statement_path, 'r' ) as statement_file:
      statement_csv = csv.reader( statement_file )

      next( statement_csv ) # skip the header
      for row in statement_csv:
        amount = USD.FromString( row[ bank_info.amount_idx ] )
        date   = datetime.strptime( row[ bank_info.date_idx ], bank_info.date_fmt )
        name   = row[ bank_info.name_idx ]
        category = categorizer.GetCategoryForName( name )
        transaction = Transaction( name, amount, date, category )
        self.transactions.append( transaction )

        if self.start_date is None or self.start_date > date:
          self.start_date = date

        if self.end_date is None or self.end_date < date:
          self.end_date = date
        
      # extract general info about statement (account, start/end date if available)
      if bank_info.date_begin_pattern is not None:
        inferred_date = bank_info.date_begin_pattern.Match( file_basename, 'start date', bank_info.bank.name )
        if inferred_date is not None:
          self.start_date = inferred_date

      if bank_info.date_end_pattern is not None:
        inferred_date = bank_info.date_end_pattern.Match( file_basename, 'end date', bank_info.bank.name )
        if inferred_date is not None:
          self.end_date = inferred_date

      if bank_info.account_id_pattern is not None:
        m = re.match( bank_info.account_id_pattern, file_basename )
        if m is not None:
          self.account_id = int( m.group(1) )

  def ResolveTransfers( self, statements ) -> None:
    for tx in [ tx for tx in self.transactions if tx.category == 'Transfer' ]:
      if re.search( r'E\-PAYMENT DISCOVER', tx.name ) is not None:
        tx.matched_transfer = True

    for stmt in statements:
      stmt_tx_amts = [ tx.amount for tx in stmt.transactions if tx.category == 'Transfer' ]
      for tx in [ tx for tx in self.transactions if tx.category == 'Transfer' ]:
        if -tx.amount in stmt_tx_amts:
          tx.matched_transfer = True
