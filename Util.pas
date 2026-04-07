unit Util;

interface

  { *** Procedimentos para serem usadas em todo os sistema ***}


  { *** Funções para serem usadas em todo os sistema ***}
  function ValorParaFloat(const V: Variant): Double;

implementation

uses
  System.Variants, System.SysUtils;

function ValorParaFloat(const V: Variant): Double;
begin
  Result := 0;

  if VarIsNull(V) or VarIsEmpty(V) then
    Exit;

  try
    if VarIsNumeric(V) then
      Result := V
    else
      Result := StrToFloat(StringReplace(VarToStr(V), '.', ',', [rfReplaceAll]));
  except
    Result := 0;
  end;
end;

end.
