import math
import re

class USD:
  dollar_amt : int
  cents_amt  : int
  incoming   : bool

  def __init__( self, dollar_amt : int, cents_amt : int, incoming : bool ) -> None:
    self.dollar_amt = dollar_amt
    self.cents_amt  = cents_amt
    self.incoming   = incoming

  def FromString( string : str ):
    new_usd = USD(0,0,False)
    if re.match( r'\-?\d+\.\d+', string ) is not None:
      amount_text = string.split('.')
      
      #subtle: we're flipping the sign here.
      # cash is represented as negative for income in csv
      # internally, negative means spent
      new_usd.dollar_amt = abs(int(amount_text[0]))
      new_usd.cents_amt  =     int(amount_text[1])
      new_usd.incoming   = math.copysign( 1, int(amount_text[0]) ) == -1

    elif re.match( r'\(?\$(\d+)\.?(\d+)?\)?', string ) is not None:
      m = re.match( r'\(?\$(\d+)\.?(\d+)?\)?', string )

      new_usd.dollar_amt = int(m.group(1))
      new_usd.cents_amt  = int(m.group(2)) if m.group(2) is not None else 0
      new_usd.incoming   = string.startswith( '(' ) == False
    return new_usd

  def __add__( self, usd ):
    int_amt       = self.dollar_amt * 100 + self.cents_amt
    int_other_amt = usd.dollar_amt * 100 + usd.cents_amt
    if self.incoming == False:
      int_amt *= -1
    if usd.incoming == False:
      int_other_amt *= -1

    total_amt = int_amt + int_other_amt
    
    return USD( math.floor( abs( total_amt / 100 ) ), abs( total_amt % 100 ), math.copysign( 1, total_amt ) == 1 )

  def __repr__( self ) -> str:
    neg : str = '' if self.incoming else '-'
    return f'{neg}${self.dollar_amt}.{self.cents_amt}'
  
  def __neg__( self ):
    return USD( self.dollar_amt, self.cents_amt, not self.incoming )
  
  def __eq__( self, other ):
    return self.dollar_amt == other.dollar_amt and self.cents_amt == other.cents_amt and self.incoming == other.incoming

  def AsExcel( self ) -> str:
    neg : str = '' if self.incoming else '-'
    return f'{neg}{self.dollar_amt}.{self.cents_amt}'