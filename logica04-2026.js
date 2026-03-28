function dispararAtualizacao(e) {
  if (!e) return;

  var planilha = e.source;
  var abaAtiva = e.range.getSheet();
  var nomeAbaAtiva = abaAtiva.getName();
  var nomesAbas = ["S.LIQ", "CLORO", "S.GRAN", "CAL"];
  
  var indiceAba = nomesAbas.indexOf(nomeAbaAtiva);
  if (indiceAba === -1) return;

  var guiaDestinoResumo = planilha.getSheetByName("S.LIQ");
  var linha = e.range.getRow();
  var col = e.range.getColumn();
  
  if (linha >= 6 && linha <= 67 && (col >= 3 && col <= 5)) {
    var dadosAba = abaAtiva.getRange("C6:E67").getValues();
    var resultadosF = [];
    var proxC = [];

    for (var i = 0; i < dadosAba.length; i++) {
      var c = (i === 0) ? dadosAba[i][0] : proxC[i - 1][0];
      var d = dadosAba[i][1] || 0;
      var eVal = dadosAba[i][2] || 0;
      var saldo = c - d + eVal;
      
      resultadosF.push([saldo]);
      if (i < dadosAba.length - 1) proxC.push([saldo]);
    }
    abaAtiva.getRange("F6:F67").setValues(resultadosF);
    abaAtiva.getRange("C7:C67").setValues(proxC);
  }  
  
  if (col === 4 || col === 5) {
    var idExterno = "1-IelB4kKlG1A71CyZ8JnDkjmsXeVBFpX8W-IAfB25_s"; 
    
    if (guiaDestinoResumo) {      
      var dadosDLocal = abaAtiva.getRange("D6:D63").getValues(); 
      var dadosELocal = abaAtiva.getRange("E6:E63").getValues(); 
      
      var somaExtSaidaD = 0;
      var somaExtEntradaE = 0;

      try {        
        var pExt = SpreadsheetApp.openById(idExterno);
        var vExt = pExt.getSheetByName(nomeAbaAtiva).getRange("D64:E67").getValues(); 
        
        for (var k = 0; k < vExt.length; k++) { 
          somaExtSaidaD += Number(vExt[k][0]) || 0;
          somaExtEntradaE += Number(vExt[k][1]) || 0;
        }
      } catch (err) {
        console.log("Erro na busca externa: " + err);
      }

      var somarIntervalo = function(dados, ini, fim) {
        var total = 0;
        for (var i = ini; i <= fim; i++) { total += Number(dados[i][0]) || 0; }
        return total;
      };
      
      var totalI = somarIntervalo(dadosDLocal, 0, 9) + somaExtSaidaD;
      var totalJ = somarIntervalo(dadosELocal, 0, 9) + somaExtEntradaE;
      
      guiaDestinoResumo.getRange(4 + indiceAba, 9).setValue(totalI);  
      guiaDestinoResumo.getRange(4 + indiceAba, 10).setValue(totalJ); 
      
      var offsets = [11, 18, 25];
      var faixas = [[10, 23], [24, 37], [38, 51]];
      
      for (var f = 0; f < offsets.length; f++) {
        guiaDestinoResumo.getRange(offsets[f] + indiceAba, 9).setValue(somarIntervalo(dadosDLocal, faixas[f][0], faixas[f][1]));
        guiaDestinoResumo.getRange(offsets[f] + indiceAba, 10).setValue(somarIntervalo(dadosELocal, faixas[f][0], faixas[f][1]));
      }
    }
  }
}
