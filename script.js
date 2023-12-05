// Google Script

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('Menu');
  menu.addItem('Pedido de transporte', 'form_pedido_transporte');
  menu.addToUi();
}

function form_pedido_transporte() {
  let emailSolicitante = CalendarApp.getId();
  const nomeSolicitante = getUserByEmail(emailSolicitante)[0]

  let html = `

    <form>
      <table>
        <tr>
          <td>
            <label for="cnpj"><b>Nome do solicitante:</b></label>
            <input type="text" id="solicitante" name="solicitante" placeholder="Informe o seu nome" value="${nomeSolicitante}">
          </td>
          <td>
            <label for="cnpj"><b>CNPJ do remetente:</b></label>
            <input type="text" id="cnpj_remetente" name="cnpj" placeholder="Informe o CNPJ" onblur="validaCNPJ()">
          </td>
          <td>
            <label for="cnpj"><b>CNPJ do destinatário:</b></label>
            <input type="text" id="cnpj_destinatario" name="cnpj" placeholder="Informe o CNPJ" onblur="validaCNPJ()">
          </td>
          <td>
            <label for="cnpj"><b>CNPJ do pagador:</b></label>
            <input type="text" id="cnpj_pagador" name="cnpj" placeholder="Informe o CNPJ" onblur="validaCNPJ()">
          </td>
        </tr>
        <tr>
          <td>
            <label><b>Pagador:</b></label><br>
          </td>

        </tr>
        <tr>
          <td>
            <div class="radio-label">
            <input type="radio" id="remetente" name="pagador" value="remetente">
            <label for="remetente">Remetente</label><br>
            </div>
          </td>
        </tr>
        <tr>
          <td>
            <div class="radio-label">
            <input type="radio" id="destinatario" name="pagador" value="destinatario">
            <label for="destinatario">Destinatário</label><br>
            </div>
          </td>
        </tr>
        <tr>
          <td>
            <div class="radio-label">
            <input type="radio" id="outro" name="pagador" value="outro">
            <label for="outro">Outro</label>
            </div>
          </td>
        </tr>
      </table>

      <br><br>

      
      <br><br>

      <input type="submit" value="Enviar">
      <br><br>

      <div id="erro"></div>

    </form>
    `

  let js =
    `
      <script>
        

        function formataCNPJ(x) {
          //Recebe um valor de CNPJ com ou sem caracteres e formata ele para "00.000.000/0001-00"

          // Remove caracteres não numéricos
          let valor = x;
          let cnpj = "";
          for (var i = 0; i < valor.length; i++) {
              if (!isNaN(parseInt(valor[i]))) {
                  cnpj += valor[i];
              }
          }

          if(cnpj !== ""){
          // Adiciona a formatação desejada manualmente
          cnpj = cnpj.substring(0, 2) + '.' + cnpj.substring(2, 5) + '.' + cnpj.substring(5, 8) + '/' + cnpj.substring(8, 12) + '-' + cnpj.substring(12, 14);
          console.log(cnpj);
          }
          return cnpj;
        }

        
        function validaCNPJ() {
          var input = document.querySelector('#cnpj_remetente');
          var erro = document.querySelector('#erro');
          var valor = input.value;

          // Remove caracteres não numéricos
          let cnpj = "";
          for (var i = 0; i < valor.length; i++) {
              if (!isNaN(parseInt(valor[i]))) {
                  cnpj += valor[i];
              }
          }

          atualizarCnpjPagador()

          if (cnpj.length !== 14) {
            erro.textContent = 'CNPJ do remetente deve conter 14 dígitos.'
            return false;
          }

          // Verifica se todos os dígitos são iguais
          if (todosDigitosIguais(cnpj)) {
            erro.textContent = 'CNPJ do remetente inválido, todos os dígitos são iguais.'
            return false;
          }

          // Calcula o primeiro dígito verificador
          var tamanho = cnpj.length - 2;
          var numeros = cnpj.substring(0, tamanho);
          var digitos = cnpj.substring(tamanho);
          var soma = 0;
          var pos = tamanho - 7;

          for (var i = tamanho; i >= 1; i--) {
              soma += numeros.charAt(tamanho - i) * pos--;
              if (pos < 2) {
                   pos = 9;
               }
          }

          var resultado = soma % 11 < 2 ? 0 : 11 - (soma % 11);

          if (resultado != digitos.charAt(0)) {
            erro.textContent = 'CNPJ do remetente inválido, primeiro dígito verificador incorreto.'
            return false;
          }

          // Calcula o segundo dígito verificador
          tamanho = tamanho + 1;
          numeros = cnpj.substring(0, tamanho);
          digitos = cnpj.substring(tamanho);
          soma = 0;
          pos = tamanho - 7;

          for (var i = tamanho; i >= 1; i--) {
            soma += numeros.charAt(tamanho - i) * pos--;
            if (pos < 2) {
                 pos = 9;
             }
          }

          resultado = soma % 11 < 2 ? 0 : 11 - (soma % 11);

          if (resultado != digitos.charAt(0)) {
            erro.textContent = 'CNPJ do remetente inválido, segundo dígito verificador incorreto.'
            return false;
          }

          erro.textContent = '';
          
          return true;
        }

        function todosDigitosIguais(cnpj) {
            for (var i = 1; i < cnpj.length; i++) {
                if (cnpj.charAt(i) !== cnpj.charAt(0)) {
                    return false;
                }
            }
            return true;
        }
    </script>
    
    <script>
    function atualizarCnpjPagador() {
      var cnpjPagador = document.getElementById("cnpj_pagador");
      var remetenteRadio = document.getElementById("remetente");
      var destinatarioRadio = document.getElementById("destinatario");
      var outroRadio = document.getElementById("outro");

      document.getElementById("cnpj_remetente").value = formataCNPJ(document.getElementById("cnpj_remetente").value)

      if (remetenteRadio.checked) {
        cnpjPagador.value = formataCNPJ(document.getElementById("cnpj_remetente").value);
        cnpjPagador.readOnly = true;
      } else if (destinatarioRadio.checked) {
        cnpjPagador.value = formataCNPJ(document.getElementById("cnpj_destinatario").value);
        cnpjPagador.readOnly = true;
      } else if (outroRadio.checked) {
        cnpjPagador.value = "";
        cnpjPagador.readOnly = false;
      }
    }
    </script>

    <script>
      // Adiciona eventos onclick aos radio buttons
      document.getElementById("remetente").onclick = atualizarCnpjPagador;
      document.getElementById("destinatario").onclick = atualizarCnpjPagador;
      document.getElementById("outro").onclick = atualizarCnpjPagador;

      // Chama a função inicialmente para configurar o estado inicial
      atualizarCnpjPagador();
    </script>
    
    `;



  let css = `
  <style>
    .center {
      text-align: center;
    }

    label {
        display: block;
        margin-top: 10px;
    }

    select, input.fornecedor {
        width: 100%;
        padding: 5px;
        margin-bottom: 10px;
    }

    #maisProdutos, #solTransp {
        padding: 10px;
        background-color: #007BFF;
        color: #fff;
        border: none;
        cursor: pointer;
    }

    #maisProdutos:hover, #solTransp:hover {
        background-color: #0056b3;
    }

    #aviso {
        font-weight: bold;
        color: red;
    }
  </style>
  
  <style>
  .radio-label {
    display: flex;
    align-items: center;
  }

  .radio-input {
    margin-right: 10px;
  }
  </style>
  
  `;

  let page = HtmlService.createHtmlOutput(html + js + css).setHeight(3000).setWidth(3000);
  SpreadsheetApp.getUi().showModalDialog(page, "Pedido de Transporte");
}
