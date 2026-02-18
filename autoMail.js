function enviarEmailComAlias(to, subject, htmlBody, fromEmail, fromName) {
  to = String(to).trim();

  const rawMessage =
  "From: " + fromName + " <" + fromEmail + ">\r\n" +
  "To: " + to + "\r\n" +
  "Reply-To: " + fromEmail + "\r\n" +
  "Subject: =?UTF-8?B?" + Utilities.base64Encode(subject, Utilities.Charset.UTF_8) + "?=\r\n" +
  "MIME-Version: 1.0\r\n" +
  "Content-Type: text/html; charset=UTF-8\r\n" +
  "Content-Transfer-Encoding: base64\r\n\r\n" +
  Utilities.base64Encode(htmlBody, Utilities.Charset.UTF_8);

  const encodedMessage = Utilities.base64EncodeWebSafe(rawMessage);

  Gmail.Users.Messages.send(
    { raw: encodedMessage },
    "me"
  );
}


function saudacaoPorHorario() {
  const hora = new Date().getHours();

  if (hora >= 5 && hora < 12) return 'Bom Dia!';
  if (hora >= 12 && hora < 19) return 'Boa Tarde!';
  return 'Boa Noite!';
}


function aoEnviarFormulario(e) {
  const sheet = e.range.getSheet();
  const linha = e.range.getRow();

  const colunaStatus = 1;
  const colunaEmail = 3;

  const rangeStatus = sheet.getRange(linha, colunaStatus);
  const cor = rangeStatus.getBackground().toLowerCase();
  const email = sheet.getRange(linha, colunaEmail).getValue();

  if (!email || !email.includes("@")) return;
  if (cor !== "#ffffff") return;

  const saudacao = saudacaoPorHorario();

  const htmlMensagem = `
    <div style="font-family: Arial, Helvetica, sans-serif; font-size: 16px; color: #000;">

      <p style="font-size: 17px; font-weight: bold; margin-bottom: 7px;">
        ${saudacao}
      </p>

      <p style="font-size: 16px; line-height: 1;">
        MENSAGEM EDITAR !!
        <br><br>

        <strong style="font-weight: bold;">
          MENSAGEM AUTOMÁTICA <----
        </strong>
      </p>


      <p>--</p>

      <!-- Assinatura -->
      <p style="font-size: 12px; line-height: 1.4; margin: 0; font-weight: bold; color: #444444
;">
        Atenciosamente,<br>
        Departamento<br>
        Empresa
      </p>

      <!-- LOGO DA EMPRESA -->
      <img
        src="LINK AQUI"
        alt="LOGO"
        style="width: 100px; height: auto; display: block;"
      />

    </div>
  `;

  enviarEmailComAlias(
    email,
    "ASSUNTO", // COLOCAR ASSUNTO DO EMAIL
    htmlMensagem,
    "EMAIL", // ALIAS -> EMAIL QUE VAI MANDAR A MENSAGEM
    "NOME" // RESPECTIVO NOME DO EMAIL
  );

  rangeStatus.setBackground("#00ffff");
}
