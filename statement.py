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

class Account:
  bank : Bank
  id   : int
  
  def __init__( self, bank : Bank, id : int = -1 ) -> None:
    self.bank = bank
    self.id   = id

  def __repr__( self ) -> str:
    acct_id_spec = f' {self.id}' if self.id != -1 else ''
    return f'{self.bank.name}{acct_id_spec}'

#------------------------------------------------------------------------------------------------
class Transaction:
  name             : str
  amount           : USD
  date             : datetime
  category         : str
  account          : Account
  matched_transfer : bool

  def __init__( self, name : str, amount : USD, account : Account, date : datetime, category : str ):
    self.name             = name
    self.amount           = amount
    self.account          = account
    self.date             = date
    self.category         = category
    self.matched_transfer = False

  def __repr__( self ) -> str:
    return f'{str(self.account)} : {self.amount} on {self.date}, {self.name}'

#------------------------------------------------------------------------------------------------
class Statement:
  account      : Account
  start_date   : datetime
  end_date     : datetime
  transactions : List[Transaction]
  
  def __init__( self, bank : Bank ):
    self.account      = Account( bank )
    self.transactions = []
    self.start_date   = None
    self.end_date     = None

  def __repr__( self ) -> str:
    return f'Statement for {str(self.account)} from {self.start_date} to {self.end_date}:\n  ' + '\n  '.join( [ str(t) for t in self.transactions] )
  
  def Read( self, statement_path : str, bank_info : BankInfo, categorizer : Categorizer ) -> None:
    file_basename = os.path.basename( statement_path )

    with open( statement_path, 'r' ) as statement_file:

      # get account info
      if bank_info.account_id_pattern is not None:
        m = re.match( bank_info.account_id_pattern, file_basename )
        if m is not None:
          self.account.id = int( m.group(1) )

      statement_csv = csv.reader( statement_file )

      next( statement_csv ) # skip the header
      for row in statement_csv:
        amount = USD.FromString( row[ bank_info.amount_idx ] )
        date   = datetime.strptime( row[ bank_info.date_idx ], bank_info.date_fmt )
        name   = row[ bank_info.name_idx ]
        category = categorizer.GetCategoryForName( name )
        transaction = Transaction( name, amount, self.account, date, category )
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

  def ResolveTransfers( self, statements ) -> None:
    for tx in [ tx for tx in self.transactions if tx.category == 'Transfer' ]:
      if re.search( r'E\-PAYMENT DISCOVER', tx.name ) is not None:
        tx.matched_transfer = True

    for stmt in statements:
      stmt_tx_amts = [ tx.amount for tx in stmt.transactions if tx.category == 'Transfer' ]
      for tx in [ tx for tx in self.transactions if tx.category == 'Transfer' ]:
        if -tx.amount in stmt_tx_amts:
          tx.matched_transfer = True
