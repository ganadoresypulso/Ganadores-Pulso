=LET(
  FechaHoy; HOY();
  Caballos; 'FORMATO DE REGISTRO.xlsx'!Tabla1[EJEMPLAR];
  Fechas; 'FORMATO DE REGISTRO.xlsx'!Tabla1[FECHA CARRERA];
  Pesos; 'FORMATO DE REGISTRO.xlsx'!Tabla1[PESO FISICO];
  CantLlegadas; CONTAR.SI.CONJUNTO(Caballos; C12; Fechas; "<"&FechaHoy; Pesos; ">0");

  SI(CantLlegadas=0; EXPANDIR(""; 5; 4; "");
    LET(
      mLlegadas; FILTRAR(ELEGIRCOLS('FORMATO DE REGISTRO.xlsx'!Tabla1[#Datos]; 58; 60; 61; 84); (Fechas<FechaHoy) * (Caballos=C12) * (ESNUMERO(Pesos)));
      mFormat; SI(CantLlegadas=1; INDICE(mLlegadas; 1; 0); mLlegadas);
      mCorte; SI(CantLlegadas>5; TOMAR(mFormat; -5); mFormat);
      FilasFaltantes; 5 - MIN(5; CantLlegadas);

      SI(FilasFaltantes=0; mCorte; APILARV(EXPANDIR(""; FilasFaltantes; 4; ""); mCorte))
    )
  )
)
