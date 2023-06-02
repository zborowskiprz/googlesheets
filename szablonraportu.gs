function nowywiersz () {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  spreadsheet.getCurrentCell().offset(0,2).activate();
  spreadsheet.getActiveRange().setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(false)
  .requireValueInList([   'Filiżanka porcelanowa',
                          'Wysyłka',
                          'Dzbanek porcelanowy',
                          'Cukiernica porcelanowa',
                          'Talerz porcelanowy',
                          'Misa porcelanowa',
                          'Wazon porcelanowy',
                          'Świecznik porcelanowy',
                          'Serwis Porcelanowy',
                          'Pojemnik porcelanowy', 
                          'Patera Porcelanowa', 
                          'Figurka porcelanowa', 
                          'Pozostała porcelana', 
                          'Wazon szklany', 
                          'Cukiernica szklana', 
                          'Pojemnik szklany', 
                          'Świecznik szklany', 
                          'Misa szklana', 
                          'Popielniczka szklana', 
                          'Karafka szklana', 
                          'Kieliszek szklany', 
                          'Zestaw szklany', 
                          'Szklana figurka',
                          'Szklanka kryształowa',
                          'Szkło pozostałe',
                          'Wyrób srebro',
                          'Biżuteria modowa',
                          'Wyrób ceramiczny',
                          'Zabytki techniki',
                          'Ozdoba drewniana',
                          'Wyrób papierowy',
                          'Szafka',
                          'Stolik',
                          'Półka',
                          'Ozdoba kamienna',
                          'Art. platerowane',
                          'Rama ozdobna',
                          'Odzież vintage',
                          'Radio',
                          'Art. metalowe',
                          'Zegarek naręczny',
                          'Zegar',
                          'Szkatułka',
                          'Lampa',
                          'Obraz',
                          'Pocztówka',
                          'Talerz szklany',
                          'Książka'], true)
  .build());
  //Cena jedn. Brutto Netto VAT
  spreadsheet.getCurrentCell().offset(0,1,1,4).activate();
  spreadsheet.getActiveRangeList().setNumberFormat('#,##0.00\\ [$zł-415]');
  spreadsheet.getCurrentCell().offset(0,5).activate();
  spreadsheet.getActiveRangeList().setNumberFormat('#,##0.00\\ [$zł-415]');
  spreadsheet.getCurrentCell().offset(0,1).activate();
  date = new Date();
  spreadsheet.getActiveRangeList().setValue(date.getFullYear()+'-'+(date.getMonth()+1)+'-'+date.getDate());
  spreadsheet.getActiveRangeList().setNumberFormat('yyyy-MM-dd')
  spreadsheet.getCurrentCell().offset(0,-8).activate();
}


function szablonraportu() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  apokalipsa();
  spreadsheet.getRange('P1').activate();
  spreadsheet.getCurrentCell().setFormulaR1C1('=ADDRESS(MATCH(MAX(C1:C1);C1:C1;0); 1;1;0)');
  spreadsheet.getRange('P2').activate();
  spreadsheet.getCurrentCell().setFormulaR1C1('=MATCH("LP";C1:C1;0)');
  spreadsheet.getRange('C1:K1').activate();
  spreadsheet.getActiveRange().setBackground('#999999');
  spreadsheet.getRange('C2').activate();
  spreadsheet.getCurrentCell().setValue('LP');
  spreadsheet.getActiveRangeList().setFontWeight('bold');
  spreadsheet.getRange('D2').activate();
  spreadsheet.getCurrentCell().setValue('Data');
  spreadsheet.getActiveRangeList().setFontWeight('bold');
  spreadsheet.getRange('E2:F32').activate();
  spreadsheet.getActiveRange().mergeAcross();
  spreadsheet.getRange('E2').activate();
  spreadsheet.getCurrentCell().setValue('Dokument');
  spreadsheet.getActiveRangeList().setFontWeight('bold');
  spreadsheet.getRange('G2').activate();
  spreadsheet.getCurrentCell().setValue('Wartość brutto');
  spreadsheet.getActiveRangeList().setFontWeight('bold');
  spreadsheet.getRange('H2').activate();
  spreadsheet.getCurrentCell().setValue('Wartość netto');
  spreadsheet.getActiveRangeList().setFontWeight('bold');
  spreadsheet.getRange('I2').activate();
  spreadsheet.getCurrentCell().setValue('VAT');
  spreadsheet.getActiveRangeList().setFontWeight('bold');
  spreadsheet.getRange('J2').activate();
  spreadsheet.getCurrentCell().setValue('Suma z raportu');
  spreadsheet.getActiveRangeList().setFontWeight('bold');
  spreadsheet.getRange('K2').activate();
  spreadsheet.getCurrentCell().setValue('Uwagi');
  spreadsheet.getActiveRangeList().setFontWeight('bold');
  spreadsheet.getRange('D3').activate();
  var date = new Date();
  date.setHours(0, 0, 0, 0);
  const m = date.getMonth();
  var month = date.getMonth();
  var lastDay = new Date(date.getFullYear(), m + 1, 0).getDate(); 
  for (var i=1; month < m+1 && i <= lastDay; i++) {
    date.setDate(i);
    spreadsheet.getActiveRangeList().setValue(date.getFullYear()+'-'+(date.getMonth()+1)+'-'+date.getDate());
    spreadsheet.getActiveRangeList().setNumberFormat('yyyy-MM-dd');
    spreadsheet.getCurrentCell().offset(0, 1, 1, 2).activate().mergeAcross();
    spreadsheet.getCurrentCell().offset(0, 2).activate();
    spreadsheet.getCurrentCell().setFormulaR1C1('=sumif(R36C11:C11;R[0]C[-3];R36C6:C6)');
    spreadsheet.getActiveRangeList().setNumberFormat('#,##0.00\\ [$zł-415]');
    spreadsheet.getCurrentCell().offset(0,1).activate();
    spreadsheet.getCurrentCell().setFormulaR1C1('=R[0]C[-1]*0,813')
    spreadsheet.getActiveRangeList().setNumberFormat('#,##0.00\\ [$zł-415]');
    spreadsheet.getCurrentCell().offset(0,1).activate();
    spreadsheet.getCurrentCell().setFormulaR1C1('=R[0]C[-2]-R[0]C[-1]')
    spreadsheet.getActiveRangeList().setNumberFormat('#,##0.00\\ [$zł-415]');
    spreadsheet.getCurrentCell().offset(0, -5).activate();
    spreadsheet.getCurrentCell().offset(1, 0).activate();
    var month=date.getMonth();
  }
  spreadsheet.getRange('C3').activate();
  spreadsheet.getCurrentCell().setFormulaR1C1('=sequence(COUNTA(R3C4:R[30]C4))');
  spreadsheet.getRange('C2:K2').activate();
  var currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getRange('C2:K33').applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  var banding = spreadsheet.getRange('C2:K33').getBandings()[0];
  banding = spreadsheet.getRange('C2:K33').getBandings()[0];
  banding.setHeaderRowColor(null)
  .setFirstRowColor('#ffffff')
  .setSecondRowColor('#b7b7b7')
  .setFooterRowColor(null);
  spreadsheet.getRange('D2').activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
  spreadsheet.getCurrentCell().offset(1, 1).activate();
  //Górna tabelka
  spreadsheet.getCurrentCell().setValue('Suma');
  spreadsheet.getActiveRangeList().setFontWeight('bold')
  spreadsheet.getCurrentCell().offset(0, 2).activate();
  spreadsheet.getActiveRangeList().setFormulaR1C1('=sum(R3C[0]:R[-1]C[0])');
  spreadsheet.getActiveRangeList().setFontWeight('bold')
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getActiveRangeList().setFormulaR1C1('=sum(R3C[0]:R[-1]C[0])');
  spreadsheet.getActiveRangeList().setFontWeight('bold')
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getActiveRangeList().setFormulaR1C1('=sum(R3C[0]:R[-1]C[0])');
  spreadsheet.getActiveRangeList().setFontWeight('bold')
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.PREVIOUS).activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.PREVIOUS).activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.PREVIOUS).activate();
  spreadsheet.getCurrentCell().offset(4, 0).activate();
  //Tabelka dolna
  //kolumna LP
  spreadsheet.getCurrentCell().setFormulaR1C1('={"LP";SEQUENCE(COUNTA(R[1]C[3]:C[3]))}');
  spreadsheet.getActiveRangeList().setFontWeight('bold');
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  //kolumna kod
  spreadsheet.getCurrentCell().setFormulaR1C1('=ARRAYFORMULA(IF(ROW(R[0]C[0]:C[0])=ROW();"Kod";IFNA(VLOOKUP(R[0]C[2]:C[2];Zmienne!R2C1:R49C2;2;0))))');
  spreadsheet.getActiveRangeList().setFontWeight('bold');
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  //kolumna ilość
  spreadsheet.getCurrentCell().setValue('Ilość');
  spreadsheet.getActiveRangeList().setFontWeight('bold');
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  //Kolumna nazwa
  spreadsheet.getCurrentCell().setValue('Nazwa');
  spreadsheet.getActiveRangeList().setFontWeight('bold');
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  //Kolumna cena jednostkowa
spreadsheet.getCurrentCell().setFormulaR1C1('=ARRAYFORMULA(IF(ROW(R[0]C[0]:C[0])=ROW();"Cena jedn.";IF(R[0]C[1]:C[1]>0;R[0]C[1]:C[1]/R[0]C[-2]:C[-2];)))');
  spreadsheet.getActiveRangeList().setFontWeight('bold');
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().setValue('Wartość brutto');
  spreadsheet.getActiveRangeList().setFontWeight('bold');
  spreadsheet.getCurrentCell().offset(-1, 0).activate();
  //podsumowanie nad tabelką
  spreadsheet.getActiveRangeList().setFormulaR1C1('=sum(R[2]C[0]:C[0])');
  spreadsheet.getActiveRangeList().setNumberFormat('#,##0.00\\ [$zł-415]');
  spreadsheet.getCurrentCell().offset(-1, 0).activate();
  spreadsheet.getCurrentCell().setValue('Suma brutto');
  spreadsheet.getActiveRangeList().setFontWeight('bold');
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().setValue('Suma netto');
  spreadsheet.getActiveRangeList().setFontWeight('bold');
  spreadsheet.getCurrentCell().offset(1, 0).activate();
  spreadsheet.getActiveRangeList().setFormulaR1C1('=sum(R[2]C[0]:C[0])');
  spreadsheet.getActiveRangeList().setNumberFormat('#,##0.00\\ [$zł-415]');
  spreadsheet.getCurrentCell().offset(1, 0).activate();
  //Wartość netto powrót do niższej tabelki
  spreadsheet.getCurrentCell().setFormulaR1C1('=ARRAYFORMULA(IF(ROW(R[0]C[0]:C[0])=ROW();"Wartość netto";IF(R[0]C[-1]:C[-1]>0;(R[0]C[-1]:C[-1]*0,813);)))');
  spreadsheet.getActiveRangeList().setFontWeight('bold');
  spreadsheet.getCurrentCell().offset(-2, 1).activate();
  //podsumowanie nad tabelką
  spreadsheet.getCurrentCell().setValue('Suma VAT');
  spreadsheet.getActiveRangeList().setFontWeight('bold');
  spreadsheet.getCurrentCell().offset(1, 0).activate();
  spreadsheet.getActiveRangeList().setFormulaR1C1('=sum(R[2]C[0]:C[0])');
  spreadsheet.getActiveRangeList().setNumberFormat('#,##0.00\\ [$zł-415]');
  spreadsheet.getCurrentCell().offset(1, 0).activate();
  //VAT
  spreadsheet.getCurrentCell().setFormulaR1C1('=ARRAYFORMULA(IF(ROW(R[0]C[0]:C[0])=ROW();"VAT";IF(R[0]C[-2]:C[-2]>0;(R[0]C[-2]:C[-2]-R[0]C[-1]:C[-1]);)))');
  spreadsheet.getActiveRangeList().setFontWeight('bold');
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  //%PIT
  spreadsheet.getCurrentCell().setValue('=ARRAYFORMULA(IF(ROW(R[0]C[0]:C[0])=ROW();"% PIT"; IF(R[0]C[-3]:C[-3]>0; "3%";)))');
  spreadsheet.getActiveRangeList().setFontWeight('bold');
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  //Wartość PIT
  spreadsheet.getCurrentCell().setValue('=ARRAYFORMULA(IF(ROW(R[0]C[0]:C[0])=ROW();"Wartość PIT";IF(R[0]C[-4]:C[-4]>0; (R[0]C[-3]:C[-3]*R[0]C[-1]:C[-1]);)))');
  spreadsheet.getActiveRangeList().setFontWeight('bold');
  //podsumowanie nad tabelką
  spreadsheet.getCurrentCell().offset(-1, 0).activate();
  spreadsheet.getActiveRangeList().setFormulaR1C1('=sum(R[2]C[0]:C[0])');
  spreadsheet.getActiveRangeList().setNumberFormat('#,##0.00\\ [$zł-415]');
  spreadsheet.getCurrentCell().offset(-1, 0).activate();
  spreadsheet.getCurrentCell().setValue('Suma PIT');
  spreadsheet.getActiveRangeList().setFontWeight('bold');
  spreadsheet.getCurrentCell().offset(2, 1).activate();
  //Data
  spreadsheet.getCurrentCell().setValue('Data');
  spreadsheet.getActiveRangeList().setFontWeight('bold');
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().setValue('Kanał');
  spreadsheet.getActiveRangeList().setFontWeight('bold');
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().setValue('Uwagi');
  spreadsheet.getActiveRangeList().setFontWeight('bold');
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().setValue('SKU');
  spreadsheet.getActiveRangeList().setFontWeight('bold');
  spreadsheet.getCurrentCell().offset(0, -13, 1, 14).activate();
  spreadsheet.getActiveRange().setBackground('#999999');
  spreadsheet.getCurrentCell().offset(1, 0, 1, 14).activate();
  spreadsheet.getActiveRange().applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  var banding = spreadsheet.getActiveRange().getBandings()[0];
  banding = spreadsheet.getActiveRange().getBandings()[0];
  banding.setHeaderRowColor(null)
  .setFirstRowColor('#ffffff')
  .setSecondRowColor('#b7b7b7')
  .setFooterRowColor(null);
  spreadsheet.getCurrentCell().offset(0, 1).activate();
nowywiersz();
};
