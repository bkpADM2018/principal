<%
Function CalcDayOfWeek(byval p_anio, byval p_mes, byval p_dia)
  '{ First test For error conditions on input values: }
  if (p_anio < 0) or (p_mes < 1) or (p_mes > 12) or (p_dia < 1) or (p_dia > 31) then
    CalcDayOfWeek = -1  '{ Return -1 to indicate an error }
  else
  	'{ Do the Zeller's Congruence calculation as Zeller himself }
 	'{ described it in "Acta Mathematica" #7, Stockhold, 1887.  }
    '{ First we separate out the p_anio and the century figures: }
    Century = int(p_anio/100)
    p_anio = p_anio MOD 100
    '{ Next we adjust the p_mes such that March remains p_mes #3, }
    '{ but that January and February are p_mess #13 and #14,     }
    '{ *but of the previous p_anio*: }
    if p_mes < 3 then
      p_mes = p_mes + 12
      if p_anio > 0 then
        p_anio = p_anio - 1 '{ The p_anio before 2000 is }
      else                  '{ 1999, not 20-1...       }
        p_anio = 99
        Century = Century - 1
	  end if
    end if

    '{ Here's Zeller's seminal black magic: }
    Holder = p_dia                        '{ Start With the day of p_mes }
    Holder = Holder + Int(((p_mes + 1) * 26)/10) '{ Calc the increment }
    Holder = Holder + p_anio              '{ Add in the p_anio }
    Holder = Holder + Int(p_anio/4)      '{ Correct For leap p_anios  }
    Holder = Holder + Int(Century/4)   '{ Correct For century p_anios }
    Holder = Holder - Century - Century '{ DON'T KNOW WHY HE DID THIS! }
    '{***********************KLUDGE ALERT!***************************}
    While (Holder < 0)                    '{ Get negative values up into }
      Holder = Holder + 7                 '{ positive territory before   }
	wend
	                                      '{ taking the MOD...         }
    Holder = Holder MOD 7                 '{ Divide by 7 but keep the  }
                                          '{ remainder rather than the }
                                          '{ quotient }
    '{***********************KLUDGE ALERT!***************************}
    '{ Here we "wrap" Saturday around to be the last day: }
    if Holder = 0 then Holder = 7

    '{ Zeller kept the Sunday = 1 origin; computer weenies prefer to }
    '{ start everything With 0, so here's a 20th century kludge:     }
    Holder = Holder - 1

    CalcDayOfWeek = Holder  '{ Return the end product! }
  end if
end Function
'*******************************************************************************
%>
