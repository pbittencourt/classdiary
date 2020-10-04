/* */

// globals
var ss = SpreadsheetApp.getActiveSpreadsheet();
var boletim = ss.getSheetByName('Boletim');
var resumo = ss.getSheetByName('Resumo');
var atividades = ss.getSheetByName('Atividades');
var planilhas = [
    '1prw9SSIDJVCw0AS1KOIyC3Qo_SrE-9r1yQQm847aAM0',
    '1iFgsb2nN0R2AifQT-uSToc3NDpX8CF4ag0HecVcgE6o',
    '1VlVQsUsp0iRHDlFe9UAXcEap3555fiDQeGmzB65ElgU',
    '1D_e_FCXOAPrsSUYQsnFdvpP4AHVVlscnExrjV66Bfko',
    '1n8rmceDJI6qMAfP3lrhiNlrIcEi-4CVIAxfArB4-Zfg',
    '1NnOID5O7JnCJ25L76ebbGbOLHPUZJCs2DRr7tvjzlXs',
    '10X-5lC4SkAtqbvQ-gj9ADKAAOnx6cmVK4dPhJNVfuq8',
    '1UiPzsx7WqCSlTiICqHxETg48SCW59oATSrU-ObqquH8',
    '1Mxnbx3pmcun9G8QtC9sEFrf7FQbhzZX-4gIHJpRD9_4',
    '1nOFvNVRAjWrqmusE6JytRLS0Zb9EbmHlglqcajeuoAA'
];

function importData() {
    /**
     * Importa planilhas 'Resumo' e 'Atividades' de todos 
     * os professores, através de chamadas IMPORTRANGE(), 
     * para permitir alterações síncronas entre os controles
     * de notas dos professores e o boletim do estudante.
     */

    // quantidade de planilhas que serão importadas
    var total = planilhas.length;

    // inicializa variável newRow, que corresponderá
    // à próxima linha vazia do 'Resumo' do boletim
    var newRow = 652
    
    // inicia o loop
    for (var i = 0; i < total; i++) {

        // abre workbook do professor
        var teacherWorkbook = SpreadsheetApp.openById(planilhas[i]);
        var teacherName = teacherWorkbook.getSheetByName('conf').getRange('B3').getValue();
        Logger.log('Inserindo dados da planilha de prof. ' + teacherName);
        
        /*******************************
         * INSERE DADOS EM RESUMO      *
         ******************************/

        // abre planilha 'Resumo' do professor
        var teacherResumo = teacherWorkbook.getSheetByName('Resumo');
      
        // verifica tamanho do array em resumo: 30 colunas (A:AD) x numRows,
        // no qual numRows é a quantidade de linhas que essa planilha possui.
        var numRows = teacherResumo.getLastRow();

        // constroi a fórmula de IMPORTRANGE
        var importRange = `=IMPORTRANGE("${planilhas[i]}"; "'Resumo'!A2:AD${numRows}")`

        // insere a fórmula na planilha Resumo do BOLETIM
        // (célula A, linha `newRow`)
        resumo.getRange(newRow, 1).setFormula(importRange);

        // avança para a próxima linha vazia do 'Resumo' de nosso boletim
        newRow += numRows + 1

        /*******************************
         * INSERE DADOS EM ATIVIDADES  *
         ******************************/

        // abre planilha 'Atividades' do professor

        // verifica tamanho do array em atividades: 21 linhas x numCols,
        // no qual numCols é a quantidade de colunas que essa planilha possui.
        // subtraímos 1 do resultado para descontar a coluna de cabeçalhos.

        // constroi a fórmula de IMPORTRANGE

        // insere a fórmula na planilha Atividades do BOLETIM
        // (célula A, linha `newRow`)

    }
}

function hideSheets() {
  /**
   * Oculta todas as planilhas,
   * exceto 'Boletim'.
   */
  
  var sheetsCount = ss.getNumSheets();
  var sheets = ss.getSheets();
  for (var i = 0; i < sheetsCount; i++){
    var sheet = sheets[i]; 
    var sheetName = sheet.getName();
    Logger.log(sheetName); 
    if (sheetName !== "Boletim") {
        Logger.log("HIDE!");
        sheet.hideSheet();
    }
  }
  
}

function showSheets() {
  /**
   * Exibe todas as planilhas do documento,
   * para debug e testes.
   */
  
  var sheetsCount = ss.getNumSheets();
  var sheets = ss.getSheets();
  for (var i = 0; i < sheetsCount; i++){
    var sheet = sheets[i]; 
    var sheetName = sheet.getName();
    Logger.log(sheetName); 
    Logger.log("SHOW!");
    sheet.showSheet();
  }
  
}

function linhas() {
  var sheet = ss.getActiveSheet();
  for (var i = 9; i < sheet.getMaxRows(); i = i + 23 ) {
    sheet.getRange(i, 1).activate();
    sheet.setRowHeight(i, 40);
    Utilities.sleep(800);
  }
}
