// Google Script

function formataCNPJ(cnpj) {
  // Remove caracteres não numéricos
  cnpj = cnpj.replace(/\D/g, '');

  // Adiciona a formatação desejada
  cnpj = cnpj.replace(/^(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})$/, '$1.$2.$3/$4-$5');
  Logger.log(cnpj)
  return cnpj;
}

//##############################################################################################################

function getUserByEmail(email) {
  //Pegar nome do usuário
  let users = SpreadsheetApp.openById('13sRUiiqjWldvEnhhYNc8DDss37xSiDCd9uX4dVs5X4g').getSheetByName('usuarios');
  let dadosUsers = users.getRange(3, 1, users.getLastRow(), 7).getValues();
  let user = dadosUsers.find(dado => dado[2] === email);
  return user ? user : '';
}

//##############################################################################################################
