/** @OnlyCurrentDoc */

/* ================================================
                  _             _            _      
   ___ ___  _ __ | |_ _ __ ___ | | ___    __| | ___ 
  / __/ _ \| '_ \| __| '__/ _ \| |/ _ \  / _` |/ _ \
 | (_| (_) | | | | |_| | | (_) | |  __/ | (_| |  __/
  \___\___/|_| |_|\__|_|  \___/|_|\___|  \__,_|\___|

     ███╗   ██╗ ██████╗ ████████╗ █████╗ ███████╗
     ████╗  ██║██╔═══██╗╚══██╔══╝██╔══██╗██╔════╝
     ██╔██╗ ██║██║   ██║   ██║   ███████║███████╗
     ██║╚██╗██║██║   ██║   ██║   ██╔══██║╚════██║
     ██║ ╚████║╚██████╔╝   ██║   ██║  ██║███████║ 
     ╚═╝  ╚═══╝ ╚═════╝    ╚═╝   ╚═╝  ╚═╝╚══════╝ 
                                            
   ================================================ */

/**
 * Autor: Pedro P. Bittencourt 
 * Email: contato@pedrobittencourt.com.br
 * Site: pedrobittencourt.com.br
 * Github: github.com/pbittencourt
 * Versão: 3.0
 */

// global variables
var ss = SpreadsheetApp.getActive();
var conf = ss.getSheetByName('conf');
var inicio = ss.getSheetByName('Início');
var turmas = ss.getSheetByName('Turmas');
var modelo = ss.getSheetByName('modelo');
var resumo = ss.getSheetByName('Resumo');
var atividades = ss.getSheetByName('Atividades');
var defaults = ['Início', 'Turmas', 'modelo', 'conf', 'Resumo', 'Atividades'];

// colour pallete: solarized theme
var base03 = '#002b36';
var base02 = '#073642';
var base01 = '#586e75';
var base00 = '#657b83';
var base0 = '#839496';
var base1 = '#93a1a1';
var base2 = '#eee8d5';
var base3 = '#fdf6e3';
var yellow = '#b58900';
var orange = '#cb4b16';
var red = '#dc322f';
var magenta = '#d33682';
var violet = '#6c71c4';
var blue = '#268bd2';
var cyan = '#2aa198';
var green = '#859900';
var color = [yellow, orange, red, magenta, violet, blue, cyan, green];

function onOpen() {
SpreadsheetApp.getUi()
    .createMenu('ARVI')
        .addItem('Atualizar lista de alunos', 'atualizaAlunos')
    .addSeparator()
        .addItem('Adicionar atividade contínua', 'addContinua')
        .addItem('Remover atividade contínua', 'remContinua')
    .addSeparator()
        .addItem('Inserir nova turma', 'addSheet')
        .addItem('Remover esta turma', 'remSheet')
    .addSeparator()
        .addItem('Reiniciar documento', 'resetAllShit')
    .addToUi();
}

function setNumberSheets() {
  /**
   * Após a primeira página do instalador.
   * Recebe nome do prof e quantidade de planilhas
   * que deseja instalar.
   */
  
  // nome do prof. e qtd de turmas a planilhas a serem copiadas
  var profName = inicio.getRange('B6').getValue();
  var classesNumber = parseInt(inicio.getRange('B10').getValue());
  
  // verifica se prof. preencheu as informações necessárias
  if (profName == 'Escreva seu nome' || classesNumber == '') {
    
    ss.toast('Você não preencheu todas as informações!');
    
  } else {
    
    // registra nome do prof. na planilha de configuração
    conf.getRange('B3').setValue(profName);
    
    // avança para a planilha de turmas
    turmas.activate();
    
    // última linha da planilha de turmas 
    // (a partir da qual novas serão inseridas)
    var lastRow = turmas.getLastRow();
    // 'modelo' a ser copiado:
    // linha 10: Turma | Disciplina
    // linha 11: campos de texto
    var model = turmas.getRange('B10:I11');
    /* Quantidade de linhas que serão inseridas:
       O 'modelo' já existente contém duas linhas, a primeira para legenda
       e a segunda para os campos de texto. Por este motivo devemos
       subtrair 1 do número de classes, multiplicando o resultado por 2. */
    var insertRows = (classesNumber - 1) * 2;
    
    /* Se o usuário optar por criar mais de uma turma, insere novas linhas. 
       Caso contrário, aproveita o modelo já existente para criar apenas
       uma planilha. */
    if (insertRows > 0) {
      
      // insere linhas
      turmas.insertRowsAfter(lastRow, insertRows);
      
      // copia o 'modelo'
      model.copyTo( turmas.getRange(lastRow + 1, 2, insertRows, 8) );
      
      // redimensiona as linhas com 'campos de texto'
      // (por questões estéticas, ficam bem com 30px de altura)
      for (var i = (lastRow + 1); i <= (lastRow + insertRows); i ++) {
        if (i % 2 == 0) {
          // 'pula' linha para redimensionar
          turmas.setRowHeight(i, 30);
        }
      }
    }
    
  }
  
}

function installSheets () {
  /**
   * Na planilha de turmas, recebe as informações das planilhas
   * que serão criadas. Para cada uma, cria uma cópia, através
   * da chamada da função makeCopy().
   **/
  
  // O intervalo contendo dados inicia na linha 10
  // e se estende à última linha da planilha.
  var range = turmas.getRange(10, 2, turmas.getLastRow() - 10, 7);
  var values = range.getValues();
  
  /* 
  Arrays para armazenar o nome da turma e a disciplina que o 
  professor ministra nesta turma. Nosso intervalo possui 8 colunas
  de largura, mescladas da seguinte maneira:
  [2 | 1 | 2 | 1 | 1]
  [turma | divisor | disciplina | divisor | cancelar] 
  */
  var turma = [];
  var disciplina = [];
  var cancelar = [];
  for (var i = 0; i < values.length; i++) {
    if (i % 2 == 0) {
      turma.push( values[i][0] );
      disciplina.push( values[i][3] );
      cancelar.push(values [i][6] );
    }
  }
  
  // Loop através de todas as turmas definidas anteriormente
  for (var j = 0; j < turma.length; j++) {
    
    // Verifica se o usuário marcou para cancelar a
    // criação de uma turma neste registro
    if (cancelar[j] == false) {
      /* DOUBLE-CHECKING: verifica se os dados não estão em branco,
       * pois os professores tendem a não selecionar o checkbox.
       * Eliminar essa opção numa versão posterior. */
      if (turma[j] != '' || disciplina[j] != '') {
        // Copia a planilha 'modelo'
        makeCopy(turma[j], disciplina[j], j);
      }
    }
    
  }
  
  /* 
  Reseta planilhas 'inicio' e 'turmas' para o estado inicial, caso o usuário 
  deseje instalar novas planilhas via menu. Para tanto, deletaremos em 'turmas'
  tudo o que houver da linha 11 para baixo, e redefiniremos o nome do professor
  e a quantidade de turmas. */
  var turmasLastRow = turmas.getLastRow();
  if (turmasLastRow > 11) {
    turmas.deleteRows(12, turmasLastRow - 11);
  }
  inicio.getRange('B6').setValue('Escreva seu nome');
  inicio.getRange('B10').clearContent();
  turmas.getRange('B3').setValue('Defina as disciplinas que você ministra em cada turma. Ao finalizar, clique no botão "Pronto" à direita! Se quiser mudar suas escolhas, retorne para a aba "Início".');
  
  // Esconde planilhas do instalador e das configurações.
  // Exibe planilha de Resumo
  inicio.hideSheet();
  turmas.hideSheet();
  conf.hideSheet();
  atividades.hideSheet();
  resumo.showSheet();
  Utilities.sleep(500);
  
  // Renomeia planilha
  ss.toast('Renomeando o documento ...');
  renameSS(disciplina);
  Utilities.sleep(500);
  
  // Compartilha com coordenação e "rapaz do TI"
  ss.toast('Compartilhando com a coordenação ...');
  shareDoc();
  Utilities.sleep(500);
  
  // Mensagem de sucesso!
  ss.toast('Controles de notas instalados com sucesso!');
  
}

function makeCopy(t, d, c) {
  /**
   * Cria uma cópia da planilha 'modelo'.
   * Parâmetros:
   *    t: turma
   *    d: disciplina
   *    c: posição no loop _ índice da planilha sendo criada
   *       (utilizado para colorir a aba!)
   *
   * Após a criação da planilha, atualizamos 'Resumo' e 'Atividades'.
   */
  
  /* -------------------------
     CRIANDO CÓPIA DA PLANILHA
     ------------------------- */
  
  // Mensagem de início
  ss.toast('Criando controle de ' + d + ' para ' + t + '...');
  Utilities.sleep(500);
  
  /* 
  Copia a planilha 'modelo', cuj nome seguirá o padrão "6A-RED".   
  Aproveitamos para definir um 'label' no padrão "6ARED", que será utilizado
  nos intervalos nomeados, ex. "Bim6ARED", "Con6ARED", etc.
  */
  
  var cod = codify(d)
  var sheetName = t.substring(0, 1) + t.substring(t.length - 1, t.length) + '-' + cod;
  var label = t.substring(0, 1) + t.substring(t.length - 1, t.length) + cod;
  
  var newSheet = modelo.copyTo(ss);
  newSheet.setName(sheetName);
  
  // Insere 'turma' em B1:C1
  newSheet.getRange('B1:C1').setValue(t);
  // Insere 'disciplina' em B2:C2
  newSheet.getRange('B2:C2').setValue(d);
  // Insere 'bônus' em L2 e cria um
  // intervalo nomeado no padrão 'Bon6ARED'
  newSheet.getRange('L2').setFormula('=MROUND(0,15 * SUM(K31:K36); 0,2)');
  ss.setNamedRange('Bon' + label, newSheet.getRange('L2'));
  
  /* 
  A coluna A é escondida do usuário. Nela, temos algumas variáveis
  de controle, para puxar dados de outras planilhas. A célula A1
  contém a turma, para puxar da planilha mestra os nomes dos
  estudantes e as notas compartilhadas. A célula A2 contém o código
  da disciplina, para puxar as notas dos simulados. A célula A3
  contém a turma, para utilizar no resumo. 
  Evitamos fazer referência aos valores da coluna B porque estes
  podem ser alterados pelo usuário, que eventualmente faz besteiras!
  */
  newSheet.getRange('A1').setValue(t);
  newSheet.getRange('A2').setValue(cod);
  newSheet.getRange('A3').setValue(d);
  
  /* 
  Insere uma cor bonitinha!
  Verifica se o 'índice no loop', passado como parâmetro, é inferior
  à quantidade de elementos no array 'color'. Em caso afirmativo,
  define este como índice do array de cores. Caso contrário, retorna
  o resto da divisão de c pelo tamanho do array.
  */
  if (c < color.length) {
    var colorIndex = c;
  } else {
    var colorIndex = c % color.length;
  }
  newSheet.setTabColor(color[colorIndex]);
  
  // Atualiza lista de alunos
  Utilities.sleep(500);
  atualizaAlunos(sheetName);
  
  /* ------------------
     ATUALIZANDO RESUMO
     ------------------ */
    
  // Inicia a partir da última linha do resumo
  var row = resumo.getLastRow() + 1;
  
  // Coluna A: turma
  var setTurma = `='${sheetName}'!$A$1`;
  resumo.getRange(row, 1, 25).setValue(setTurma);
  
  // Coluna B: disciplina
  var setDisciplina = `='${sheetName}'!$A$3`;
  resumo.getRange(row, 2, 25).setValue(setDisciplina);
  
  // Colunas C:D: dados dos estudantes
  var setEstudantes = `={'${sheetName}'!B4:C28}`;
  resumo.getRange(row, 3).setValue(setEstudantes);
  
  // Colunas E:I: notas dos estudantes
  var setNotas = `={'${sheetName}'!Z4:AD28}`;
  resumo.getRange(row, 5).setValue(setNotas);
  
  // Intervalos nomeados [E, F, G, H, I]
  // Modelo:
  // Bim6AGMT | Con6AGMT | Com6AGMT | Dis6AGMT | Med6AGMT
  
  // E: Bimestral
  var bimRange = resumo.getRange(row, 5, 25);
  ss.setNamedRange('Bim' + label, bimRange);
  
  // F: Contínua
  var conRange = resumo.getRange(row, 6, 25);
  ss.setNamedRange('Con' + label, conRange);
  
  // G: Complementar
  var comRange = resumo.getRange(row, 7, 25);
  ss.setNamedRange('Com' + label, comRange);
  
  // H: Disciplinar
  var disRange = resumo.getRange(row, 8, 25);
  ss.setNamedRange('Dis' + label, disRange);
  
  // I: Média
  var medRange = resumo.getRange(row, 9, 25);
  ss.setNamedRange('Med' + label, medRange);
  
  // Coluna J: situações dos estudantes
  var cell = 'I' + row;
  var setSituacao = `=IF( ISTEXT(${cell}); "..."; IF( ${cell} <= 6; "rec"; "ok"))`;
  resumo.getRange(row, 10).setFormula(setSituacao).activate();
  var setSituacaoRange = resumo.getRange(row, 10, 25).getA1Notation();
  resumo.getActiveRange().autoFill(resumo.getRange(setSituacaoRange), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  
  // Insere intervalo nomeado com todas as notas contínuas.
  // Esse intervalo segue o padrão 'Continuas6ARED' e é
  // referenciado na planilha resumo, para constar nos
  // boletins individuais.
  
  var allConName = 'Continuas' + label;
  var allConRange = newSheet.getRange('D4:I28');
  ss.setNamedRange(allConName, allConRange);
  
  // Colunas K-> AD: notas contínuas dos estudantes.
  resumo.getRange(row, 11).setValue(`={${allConName}}`);
  
  Utilities.sleep(500);
  
  /* ----------------------
     ATUALIZANDO ATIVIDADES
     ---------------------- 
     A planilha de atividades contém todas as contínuas
     solicitadas pelo professor em todas as turmas/disciplinas.
     Utilizado no boletim individual. */
  
  // Nova coluna, a partir da qual se inicia a atualização
  var newCol = atividades.getLastColumn() + 1;
  
  // Insere turma na linha 1
  atividades.getRange(1, newCol).setValue(sheetName);
  
  // Insere fórmula na linha 2
  var ativFormula = `=JOIN("|"; '${sheetName}'!E31; '${sheetName}'!K31; '${sheetName}'!L31)`;
  atividades.getRange(2, newCol).setFormula(ativFormula).activate();
  
  // Copia fórmula até a linha 21
  var ativRange = atividades.getRange(2, newCol, 20).getA1Notation();
  atividades.getActiveRange().autoFill(atividades.getRange(ativRange), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  
  Utilities.sleep(500);
  
  /* -------------
     FIM DA FUNÇÃO
     ------------- */
  
  // Mensagem de sucesso!
  newSheet.activate();
  newSheet.showSheet();
  ss.toast('Criado controle de ' + d + ' para ' + t + '!');
  Utilities.sleep(500);
  
}

function codify(d) {
    /**
     * Cria um código de três letras a partir do
     * nome da disciplina. O switch é necessário
     * para lidar com nomes que contém acentos e
     * para diferenciar [GEO]metria de [GEO]grafia.
     */
    switch (d) {
        case 'Ciências':
            var cod = 'CIE';
        break;
        case 'Física':
            var cod = 'FIS';
        break;
        case 'Geometria':
            var cod = 'GMT';
        break;
        case 'Química':
            var cod = 'QUI';
        break;
        default:
            var cod = d.substring(0, 3).toUpperCase();
    }
    return cod;
}

function addSheet() {
  /**
   * Adiciona nova turma a partir do menu.
   */
  
  // verifica se o usuário quer adicionar uma nova planilha
  var ui = SpreadsheetApp.getUi();
  var msg = 'Quer adicionar mais uma planilha de controle de notas?';
  var response = ui.alert('Adicionando uma nova planilha de controle', msg, ui.ButtonSet.YES_NO);
    
  if (response == ui.Button.YES) {
    // ativa planilha de criação de turmas
    turmas.activate();
  
    // muda o texto explicativo, tendo em vista que
    // agora só é permitida a criação de *uma* turma
    var texto = 'Defina a turma e a disciplina que você deseja criar. Ao finalizar, clique no botão "Pronto!" à direita.';
    turmas.getRange('B3').setValue(texto);
  }
}

function remSheet() {
  /**
   * Remove planilha a partir do menu.
   */
  
  var sheet = ss.getActiveSheet();
  var sheetName = sheet.getName();
  var turma = sheet.getRange('A1').getValue();
  var disciplina = sheet.getRange('A3').getValue();
  
  // exibe mensagem quando a planilha faz parte de 'defaults'
  // e, portanto, não pode ser excluída!
  if (defaults.indexOf(sheetName) != -1) {
    ss.toast('Não é possível excluir esta planilha!');
  } else {
    // verifica se o usuário quer remover a planilha
    var ui = SpreadsheetApp.getUi();
    var msg = 'Tem certeza de que deseja remover essa planilha de controle de notas? Este processo é irreversível!';
    var response = ui.alert('Remover controle de notas de '+turma+', '+disciplina, msg, ui.ButtonSet.YES_NO);
  
    if (response == ui.Button.YES) {
      
      ss.toast('Removendo planilha para\n' + turma + ', ' + disciplina + ' ...');
      
      //////////////////
      // ATUALIZA RESUMO
      //////////////////
      
      resumo.activate();
      var numRows = resumo.getLastRow();
      var range = resumo.getRange(2, 1, numRows-1, 2).getValues();
      var found = false;
      var start = 0;
      
      for (var i = 0; i < range.length; i++) {
        if (range[i][0] == turma && range[i][1] == disciplina && !found) {
          start = i + 2;
          found = true;
        } 
      }
      Utilities.sleep(1000);
      resumo.deleteRows(start, 25);
      Utilities.sleep(1000);
      
      //////////////////////
      // ATUALIZA ATIVIDADES
      //////////////////////
      
      var numCols = atividades.getLastColumn();
      var range = atividades.getRange(1, 1, 1, numCols);
      var rangeValues = range.getValues();
      var header = rangeValues[0];

      for (var i = 0; i < header.length; i++) {
        if (header[i] == sheetName) {
          var deleteCol = i + 1;
        }
      }
      Utilities.sleep(1000);
      atividades.deleteColumn(deleteCol);
      Utilities.sleep(1000);
      
      //////////////////
      // DELETA PLANILHA
      //////////////////
      
      ss.deleteSheet(sheet);
      Utilities.sleep(2000);
      
      ss.toast('Controle de notas para ' + turma + ', ' + disciplina + ' excluído com sucesso!');
      
    }
  }
}

function renameSS(disciplina) {
  /**
   * Renomeia o documento, após a instalação, de acordo
   * com o padrão "Pedro FIS-GMT — 3º bim/2020".
   * Recebe como parâmetro a lista com as disciplinas
   * ministradas e puxa de 'conf' nome do prof, ano e bimestre.
   */
  
  var unique = disciplina.filter(function(elem, index, self) {
    return index === self.indexOf(elem);
  });
  
  var professor = conf.getRange('B3').getValue();
  var periodo = conf.getRange('B4').getValue();
  var ano = conf.getRange('B5').getValue();
  
  var disciplinas = '';
  for (var i = 0; i < unique.length; i++) {
    
    if (i > 0) {
      disciplinas += '-';
    } else {
      disciplinas += ' ';
    }
    
    //switch
    switch (unique[i]) {
      
      case 'Ciências':
        var cod = 'CIE';
        break;
      
      case 'Física':
        var cod = 'FIS';
        break;
        
      case 'Geometria':
        var cod = 'GMT';
        break;
      
      case 'Química':
        var cod = 'QUI';
        break;
      
      default:
        var cod = unique[i].substring(0, 3).toUpperCase();
      
    }
    disciplinas += cod;
    
  }
  
  var firstName = professor.split(' ');
  var bimestre = periodo.substring(0, 6);
  ss.rename(`${firstName[0]}${disciplinas} — ${bimestre}\/${ano}`);
}

function addContinua() {
  /**
   * Adiciona um item à relação de atividades
   * que compõem a 'Avaliação Contínua'.
   */
  
  // planilha atual
  var sheet = SpreadsheetApp.getActiveSheet();
  var label = sheet.getName().replace('-', '');
  
  // Se essa planilha estiver no defaults,
  // retorne sem executar!
  if (defaults.indexOf(sheet.getName()) == -1) {
    ss.toast('Não é possível adicionar atividade contínua nesta planilha!');
    return
  }
  
  // na última linha da planilha se localiza a atual última atividade
  var lastRow = sheet.getLastRow();
  // 'nome' da última atividade (Cnº)  
  var lastActivity = sheet.getRange(lastRow, 4);
  // número da última atividade
  var number = lastActivity.getValue().replace('C','');
  // número da próxima atividade
  var newNumber = parseInt(number) + 1;
  
  // insere nova coluna na planilha
  sheet.insertColumnBefore(3 + newNumber);
  // numera a atividade
  sheet.getRange(3, 3 + newNumber).setValue('C' + newNumber);
  // ativa a nova coluna
  var newColumn = sheet.getRange(3, 3 + newNumber, 26).activate();
  // define as bordas
  newColumn.setBorder(true, true, true, null, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  
  // insere nova linha na planilha
  sheet.insertRowAfter(lastRow);
  // numera a atividade
  sheet.getRange(lastRow + 1, 4).setValue('C' + newNumber);
  // copia a formatação
  var lastLine = sheet.getRange(lastRow, 5, 1, 11).getA1Notation();
  var lastLines = sheet.getRange(lastRow, 5, 2, 11).getA1Notation();
  sheet.getRange(lastLine).activate();
  sheet.getActiveRange().autoFill(sheet.getRange(lastLines), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  sheet.getRange(lastLines).activate();
  // limpa o conteúdo após copiar a formatação
  sheet.getRange(lastRow+1, 5, 1, 11).clearContent().activate();
  // foca no campo de descrição da atividade recém criada
  sheet.getRange(lastRow+1, 5).activate();
  
  // atualiza transposição de valores
  var rowsRange = sheet.getRange(lastRow - number + 1, 4 + newNumber + 1, newNumber).getA1Notation();
  sheet.getRange('D2').setFormula('=TRANSPOSE(' + rowsRange + ')');
  
  // atualiza bônus
  var bonusLabel = 'Bon' + label;
  ss.getRangeByName(bonusLabel).setFormula(`=MROUND(0,15 * SUM(${rowsRange}); 0,2)`);
  
  // atualiza namedRange com todas as contínuas
  var continuasRange = sheet.getRange(4, 4, 25, newNumber);
  ss.setNamedRange('Continuas' + label, continuasRange);
  
  // formatação condicional (nota vermelha)
  var col = newNumber+3;
  var range = sheet.getRange(4, col, 25);
  var rangeA1 = range.getA1Notation();
  var letter = columnToLetter(col);
  var formula = `=(\{${rangeA1}\}) < ( 0,6 * ${letter}$2 )`;
    
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(formula)
    .setFontColor('#ff0000')
    .setRanges([range])
    .build();
  var rules = sheet.getConditionalFormatRules();
  rules.push(rule);
  sheet.setConditionalFormatRules(rules);
  
  // validação de dados (nota abaixo do limite)
  var formula = `=OR( \{${rangeA1}\} <= ${letter}$2; ISTEXT(\{${rangeA1}\}) )`;
  var rule = SpreadsheetApp.newDataValidation()
    .requireFormulaSatisfied(formula)
    .setAllowInvalid(false)
    .setHelpText('Insira um valor que não exceda o estipulado para essa atividade!')
    .build();
  range.setDataValidation(rule);
  
  // largura da coluna
  sheet.setColumnWidth(col, 35);
  
  ss.toast('Atividade adicionada com sucesso!');
  
}

function remContinua() {
  /**
   * Remove o último item da lista de atividades
   * que compõem a Avaliação Contínua. Não é executado
   * caso houver apenas seis atividades (qtd mínima).
   */
  
  // pega a planilha atual e a interface gráfica
  var sheet = ss.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  var label = sheet.getName().replace('-', '');
  
  // Se essa planilha estiver no defaults,
  // retorne sem executar!
  if (defaults.indexOf(sheet.getName()) == -1) {
    ss.toast('Não é possível remover atividade contínua desta planilha!');
    return
  }
  
  // 'cabeçalho' da lista de atividades
  // utilizado pra saber quantas existem no momento
  var first = 30;
  // última linha de atividades
  var last = sheet.getLastRow();
  // quantidade de atividades
  var count = last - first;
  // coluna a ser deletada
  var columnPosition = count + 3;
  
  // se houver apenas 6 atividades, não deletaremos!
  if (count <= 6) {
    ss.toast('O número mínimo de atividades foi atingido!');
  } else {
    
    // informa o usuário pra não fazer caquinha
    var msg = 'Você tem certeza de que deseja excluir a atividade "C' + count + '" da lista? Este processo é irreversível.';
    var response = ui.alert('Deletando a última atividade da lista', msg, ui.ButtonSet.YES_NO);
    
    if (response == ui.Button.YES) {
      // remove coluna e linha
      sheet.deleteColumn(columnPosition);
      sheet.deleteRow(last);
  
      // atualiza transposição de dados
      var rowsRange = sheet.getRange(first + 1, 4 + count, count - 1).getA1Notation();
      sheet.getRange('D2').setFormula('=TRANSPOSE(' + rowsRange + ')');
      
      // atualiza namedRange com todas as contínuas
      var continuasRange = sheet.getRange(4, 4, 25, count - 1);
      ss.setNamedRange('Continuas' + label, continuasRange);
    }
    
  }
  
}

function shareDoc() {
  /**
   * Compartilha spreadsheet com coordenação e 'rapaz do TI'
   */
  
  //verifica quem são os editores do documento
  var editors = ss.getEditors(); 
  
  //lista de compartilhamento
  var shares = ['dora.bertotti@escolartedeviver.com.br', 'pedro.bittencourt@escolartedeviver.com.br'];
  var cont = 0;
  
  //percorre a lista de editores
  for (var i = 0; i < editors.length; i++) {
    var editor = editors[i];
    
    //se o editor atual do documento pertencer à lista de compartilhamentos,
    //acrescenta 1 à contagem
    if (shares.indexOf(editor) != 1) {
      cont++;
    }
  }
  
  //se a contagem de editores pertencente à lista for inferior à quantidade
  //de editores do compartilhamento, adiciona-os todos como editores do documento.
  //aproveita para conceder permissão a todos com o link
  if (cont < shares.length) {
    ss.addEditors(shares);
    var id = ss.getId();
    var doc = DriveApp.getFileById(id);
    doc.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
  }
  
  ss.toast('Lista de editores atualizada com sucesso!');
  
};

function atualizaAlunos(planilha) {
  /**
   * Atualiza relação de alunos, exibindo linhas ocultas, 
   * correspondentes a estudantes recém-matriculados.
   */
  
  // Se essa planilha estiver no defaults,
  // retorne sem executar!
  if (defaults.indexOf(planilha) != -1) {
    ss.toast('Não há lista de alunos nesta planilha!');
    return
  }
  
  //abre o documento
  var spreadsheet = SpreadsheetApp.getActive();
  if (planilha == null) {
    //se não for passado parâmetro 'planilha', seleciona a planilha ativa
    var sheet = spreadsheet.getActiveSheet();
  } else {
    //caso contrário, selecione a planilha passada como parâmetro
    var sheet = spreadsheet.getSheetByName(planilha)
  }
  
  //pega o intervalo que contém nomes dos alunos (25 linhas ao todo)
  var lista = sheet.getRange('C4:C28');
  lista.activate();
  //exibe todas as linhas
  sheet.showRows(spreadsheet.getActiveRange().getRow(), spreadsheet.getActiveRange().getNumRows());
  
  //posição do primeiro estudante
  var pos = 4;
  
  //armazena todos os valores num array
  var nomes = lista.getValues();
  
  for (var i = 0; i < nomes.length; i++) {
    //nome do aluno
    var nome = nomes[i][0];
    //quantidade de letras no nome
    var letras = nome.length;
    
    //verifica se é um nome, a partir da contagem de letras
    if (letras > 0) {
      //em caso positivo, acrescenta 1 à contagem de linhas (a partir da 1ª)
      pos++;
    } //se não houver aluno, para a contagem de linhas aqui
  }
  
  //pega o intervalo do último aluno até a linha 28 (última linha da lista de alunos)
  sheet.getRange(pos, 1, (28-pos+1)).activate();
  
  //finalmente, oculta essas linhas selecionadas
  spreadsheet.getActiveSheet().hideRows(spreadsheet.getActiveRange().getRow(), spreadsheet.getActiveRange().getNumRows());
  
  //coloca o cursor no título da planilha só pra ficar bonitin
  spreadsheet.getRange('A1').activate();
  
  ss.toast('Lista de estudantes atualizada com sucesso!');
  
}

function resetAllShit() {
    /**
    * Deleta todas as turmas que foram criadas
    * e retorna as planilhas de instalação e
    * configuração ao estado inicial.
    */

    var ui = SpreadsheetApp.getUi();
    var sheets = ss.getSheets();
    var count = sheets.length;

    // informa o usuário pra não fazer caquinha
    var msg = 'Você tem certeza de que deseja excluir todos os controles de notas deste documento e retorná-lo ao estado inicial? Este processo é irreversível.';
    var response = ui.alert('Deletando os controles de notas', msg, ui.ButtonSet.YES_NO);

    if (response == ui.Button.YES) {

        for (i = 0; i < count; i++) {
            var name = sheets[i].getName();

            // Verifica se o nome da planilha atual consta em 'DEFAULTS'
            if (defaults.indexOf(name) == -1) {  // não consta!

                // Deleta intervalos nomeados.
                // Precisamos deletá-los *antes* de excluir a planilha porque o Sheets
                // não deleta intervalos inválidos. Ou seja: se houver '#REF!', o nome
                // fica ali *para sempre*!
                var namedRanges = ss.getNamedRanges();
                var keepNamedRanges = ['LinkMestra', 'LinkSimulados', 'Professor', 'Periodo', 'Ano'];
                for (var i = 0; i < namedRanges.length; i++ ) {
                    var thisNamed = namedRanges[i].getName();
                    if (keepNamedRanges.indexOf(thisNamed) == -1) {
                        Logger.log('Deletarei ' + thisNamed);
                        namedRanges[i].remove();
                    }
                }
                Logger.log('Deletarei a planilha ' + name);

                // Deleta a planilha
                ss.deleteSheet(sheets[i]);

            } else {  // consta!

                //do nothing! grab yourself some coffee (:
                Logger.log('Manterei a planilha ' + name);

            }
        }
          
        // Reseta 'Turmas' para o estado inicial,
        // removendo o que houver da linha 11 para baixo
        // (se houver) e limpando a caixa de cancelamento
        var turmasLastRow = turmas.getLastRow();
        if (turmasLastRow > 11) {
            turmas.deleteRows(12, turmasLastRow - 11);
        }
        turmas.getRange('B10').clearContent();
        turmas.getRange('E10').clearContent();
        turmas.getRange('H10').setValue(false);

        // Reseta 'Início'
        inicio.getRange('B6').setValue('Escreva seu nome');
        inicio.getRange('B10').clearContent();

        // Reseta 'Resumo', limpando o conteúdo
        // da linha 1 para baixo (se houver)
        var resumoLastRow = resumo.getLastRow();
        if (resumoLastRow > 1) {
            resumo.getRange(2, 1, resumoLastRow - 1, 30).clearContent();
        }

        // Reseta 'Atividades', limpando o conteúdo
        // da coluna 2 para a direita
        var atividadesMaxCol = atividades.getMaxColumns()
        var atividadesMaxRow = atividades.getMaxRows();
        atividades.getRange(1, 2, atividadesMaxRow, atividadesMaxCol - 1).clearContent();

        // Exibe 'Inicio' e oculta o restante
        inicio.showSheet();
        turmas.hideSheet();
        conf.hideSheet();
        resumo.hideSheet();
        modelo.hideSheet();
        atividades.hideSheet()

    }
  
}

function columnToLetter(column) {
  /**
   * Column to Letter
   * from StackOverflow: http://stackoverflow.com/questions/21229180/convert-column-index-into-corresponding-column-letter
   */
  
  var temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

/**
 * Daqui pra baixo, funções para teste, apenas.
 * Mantenho aqui por questões de reutilização,
 * não sendo utilizadas no programa, em si.
 */

function validas() {
  var sheet = ss.getActiveSheet();
  for (var i = 4; i <= 9; i++) {
    var range = sheet.getRange(4, i, 25);
    var rangeA1 = range.getA1Notation();
    var letter = columnToLetter(i);
    var formula = `=OR( \{${rangeA1}\} <= ${letter}$2; ISTEXT(\{${rangeA1}\}) )`;
    
    var rule = SpreadsheetApp.newDataValidation()
    .requireFormulaSatisfied(formula)
    .setAllowInvalid(false)
    .setHelpText('Insira um valor que não exceda o estipulado para essa atividade!')
    .build();
    range.setDataValidation(rule);
  }
}

function formatas() {
  var sheet = ss.getActiveSheet();
  for (var i = 4; i <= 9; i++) {
    var range = sheet.getRange(4, i, 25);
    var rangeA1 = range.getA1Notation();
    var letter = columnToLetter(i);
    var formula = `=(\{${rangeA1}\}) < ( 0,6 * ${letter}$2 )`;
    
    var rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(formula)
      .setFontColor('#ff0000')
      .setRanges([range])
      .build();
    var rules = sheet.getConditionalFormatRules();
    rules.push(rule);
    sheet.setConditionalFormatRules(rules);
  }
}

function nomeas() {
  var disciplina = ['Geografia', 'Geografia', 'Geografia'];
  
  var unique = disciplina.filter(function(elem, index, self) {
    return index === self.indexOf(elem);
  });
  
  var professor = conf.getRange('B3').getValue();
  var periodo = conf.getRange('B4').getValue();
  var ano = conf.getRange('B5').getValue();
  
  var disciplinas = '';
  for (var i = 0; i < unique.length; i++) {
    
    if (i > 0) {
      disciplinas += '-';
    } else {
      disciplinas += ' ';
    }
    
    //switch
    switch (unique[i]) {
      
      case 'Ciências':
        var cod = 'CIE';
        break;
      
      case 'Física':
        var cod = 'FIS';
        break;
        
      case 'Geometria':
        var cod = 'GMT';
        break;
      
      case 'Química':
        var cod = 'QUI';
        break;
      
      default:
        var cod = unique[i].substring(0, 3).toUpperCase();
      
    }
    disciplinas += cod;
    
  }
  
  var firstName = professor.split(' ');
  var bimestre = periodo.substring(0, 6);
  //ss.rename(`${firstName[0]}${disciplinas} — ${bimestre}\/${ano}`);
  Logger.log(`${firstName[0]}${disciplinas} — ${bimestre}\/${ano}`);
  
}
