//----------------------------------------------
// Contar as células com cores: a 1ª coluna é que contém on nomes:
// Retorna um ARRAY[]
//----------------------------------------------
function countcoloredcells(countRange, countRef, colorRef) {

  var activeRange = SpreadsheetApp.getActiveRange();
  var activeSheet = activeRange.getSheet();
  var formula = activeRange.getFormula();
  
  var f = formula.match(/\((.*)\)/).pop().split(";").map( value => value.trim() ) // extrair os 3 parametros da função...
  var _countRange = f[0]
  var _countRef   = f[1]
  var _colorRef   = f[2]

  var cores = activeSheet.getRange(_colorRef).getBackgrounds();  // extrair as cores...
  
  var bgValues = activeSheet.getRange(_countRange).getValues()
  var bg       = activeSheet.getRange(_countRange).getBackgrounds().map(function(value, index) {
    value[0] = bgValues[index][0]
    return value
  });

  var lista = activeSheet.getRange(_countRef).getValues().map(function(value) {
    var b = new Array( cores[0].length + 1 ).fill(0)
    b[0] = value[0]
    return b
  });

  for(var j = 0; j < lista.length; j++) {  // lista fixa dos nomes...
    for ( var k = 0; k < bg.length; k++) {
      if ( lista[j][0] == bg[k][0] ) {
        for (z = 1; z < bg[k].length; z++){
          for (c = 0; c < cores[0].length; c++){
            if (bg[k][z] == cores[0][c]) {
              ++lista[j][c+1]
            } 
          }
        }
      }
    }
  }

  return lista;

};
//----------------------------------------------------------------


//-------------------------------------------------------------------------------------
// Contar as células com cores: os nomes estão dentro das células coloridas: 
// Retorna um ARRAY[]
//-------------------------------------------------------------------------------------
function countcoloredcells2(countRange, countRef, colorRef) {

  var activeRange = SpreadsheetApp.getActiveRange();
  var activeSheet = activeRange.getSheet();
  var formula = activeRange.getFormula();
  
  // extrair os 3 parametros da fórmula:
  var f = formula.match(/\((.*)\)/).pop().split(";").map( value => value.trim() ) 
  var _countRange = f[0]
  var _countRef   = f[1]
  var _colorRef   = f[2]

  var cores = activeSheet.getRange(_colorRef).getBackgrounds();  // cores...

  // criar o array bi-dimensional com os nomes e as cores:
  var lista = activeSheet.getRange(_countRef).getValues().map(function(value) {
    var b = new Array(cores[0].length + 1).fill(0)
    b[0] = value[0]
    return b
  });

  // somente a coluna[0], com os nomes. Mais simples no IndexOf!
  var listaUnica = lista.map(function(values) { return values[0] });

  var bgValues = activeSheet.getRange(_countRange).getValues()
  var bgCores  = activeSheet.getRange(_countRange).getBackgrounds()

  for (i = 0; i < bgValues.length; i++) {
    for (j = 0; j < bgValues[i].length; j++) {
      if (bgValues[i][j] != "" ) {
        var pos = listaUnica.indexOf(bgValues[i][j])
        if (pos > -1) {
          for (c = 0; c < cores[0].length; c++){
            if (bgCores[i][j] == cores[0][c]) {
              ++lista[pos][c+1]
            } 
          }
        }
      }
    }
  }

  return lista;

}
//.....................
function findFormula() {
  // PESQUISAR ONDE TEM AS FORMULAS: [ countcoloredcells, countcoloredcells2]
  // pesquisa apenas onde tem as funções em uso. Agiliza o Refresh;

  SpreadsheetApp.getActive().toast('Recontando as cores...', 'RECALCULANDO', 3);

  var sheet = SpreadsheetApp.getActiveSheet()
  var formulas = sheet.getDataRange().getFormulas()

  var a1 = "countcoloredcells" == "countcoloredcells2"
  var a2 = "countcoloredcells2" == "countcoloredcells"

  formulas.map( function(values, l) {
    values.map( function(val, c) {
      if (val != "") {
        if ( val.includes("countcoloredcells") ) {
          var cellF = sheet.getRange(l+1,c+1)
          refreshSheet(cellF)
        }
      }
    })
  })
}
//---------
function refreshSheet(cellFormula) {
  var formula = cellFormula.getFormula();
  cellFormula.setFormula("");
  SpreadsheetApp.flush();
  cellFormula.setFormula(formula);
}


