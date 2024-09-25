import csv
import os
import re
import yaml
from datetime import datetime
from typing import List, Dict, Pattern
from bank   import Bank, BankInfo, BankType
from currency import USD

#------------------------------------------------------------------------------------------------
class Categorizer:
  categories : Dict[ str, List[Pattern] ]

  def __init__( self, path_to_config ) -> None:
    self.categories = {}
    with open( path_to_config, 'r' ) as config_file:
      config_yaml = yaml.safe_load( config_file )
      for category in config_yaml.keys():
        if config_yaml[ category ] is not None:
          self.categories[ category ] = [ re.compile( pattern ) for pattern in config_yaml[ category ] ]
        else:
          self.categories[ category ] = []
    self.categories[ 'Unknown' ] = [ r'^$' ]

  def GetCategoryForName( self, name : str ) -> str:
    for cat, patterns in self.categories.items():
      for pattern in patterns:
        if re.search( pattern, name ) is not None:
          return cat
        
    return 'Unknown'

#------------------------------------------------------------------------------------------------
class Account:
  bank_info : BankInfo
  id        : int
  
  def __init__( self, bank_info : BankInfo, id : int = -1 ) -> None:
    self.bank_info = bank_info
    self.id        = id

  def __repr__( self ) -> str:
    acct_id_spec = f' {self.id}' if self.id != -1 else ''
    return f'{self.bank_info.bank.name}{acct_id_spec}'

#------------------------------------------------------------------------------------------------
class Transaction:
  name             : str
  amount           : USD
  date             : datetime
  category         : str
  account          : Account
  matched_transfer : bool
  is_debt          : bool

  def __init__( self, name : str, amount : USD, account : Account, date : datetime, category : str, is_debt : bool ):
    self.name             = name
    self.amount           = amount
    self.account          = account
    self.date             = date
    self.category         = category
    self.matched_transfer = False
    self.is_debt          = is_debt

  def __repr__( self ) -> str:
    return f'{str(self.account)} : {self.amount} on {self.date}, {self.name}'

#------------------------------------------------------------------------------------------------
class Statement:
  bank_info    : BankInfo
  account      : Account
  start_date   : datetime
  end_date     : datetime
  transactions : List[Transaction]
  
  def __init__( self, bank_info : BankInfo ):
    self.bank_info    = bank_info
    self.account      = Account( bank_info )
    self.transactions = []
    self.start_date   = None
    self.end_date     = None

  def __repr__( self ) -> str:
    return f'Statement for {str(self.account)} from {self.start_date} to {self.end_date}:\n  ' + '\n  '.join( [ str(t) for t in self.transactions] )
  
  def Read( self, statement_path : str, categorizer : Categorizer ) -> None:
    file_basename = os.path.basename( statement_path )

    with open( statement_path, 'r' ) as statement_file:

      # get account info
      if self.bank_info.account_id_pattern is not None:
        m = re.match( self.bank_info.account_id_pattern, file_basename )
        if m is not None:
          self.account.id = int( m.group(1) )

      statement_csv = csv.reader( statement_file )

      next( statement_csv ) # skip the header
      for row in statement_csv:
        amount   = USD.FromString( row[ self.bank_info.amount_idx ] )
        date     = datetime.strptime( row[ self.bank_info.date_idx ], self.bank_info.date_fmt )
        name     = row[ self.bank_info.name_idx ]
        category = categorizer.GetCategoryForName( name )
        is_debt  = self.account.bank_info.type == BankType.CreditCard or category == 'Loans'
        transaction = Transaction( name, amount, self.account, date, category, is_debt )
        self.transactions.append( transaction )

        if self.start_date is None or self.start_date > date:
          self.start_date = date

        if self.end_date is None or self.end_date < date:
          self.end_date = date
        
      # extract general info about statement (account, start/end date if available)
      if self.bank_info.date_begin_pattern is not None:
        inferred_date = self.bank_info.date_begin_pattern.Match( file_basename, 'start date', self.bank_info.bank.name )
        if inferred_date is not None:
          self.start_date = inferred_date

      if self.bank_info.date_end_pattern is not None:
        inferred_date = self.bank_info.date_end_pattern.Match( file_basename, 'end date', self.bank_info.bank.name )
        if inferred_date is not None:
          self.end_date = inferred_date

  def ResolveTransfers( self, statements ) -> None:
    for stmt in statements:
      stmt_tx_amts = [ tx.amount for tx in stmt.transactions if tx.category == 'Transfer' or tx.category == 'Credit Card Payment' ]
      for tx in [ tx for tx in self.transactions if tx.category == 'Transfer' or tx.category == 'Credit Card Payment' ]:
        if -tx.amount in stmt_tx_amts:
          tx.matched_transfer = True
