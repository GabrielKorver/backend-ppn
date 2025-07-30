import express from 'express';
import fetch from 'node-fetch';
import xlsx from 'xlsx';

const app = express();
const PORT = process.env.PORT || 3000;

// Configurações dos Correios
const username = "buy2buy77";
const password = "eg1PQhetrvClLeFdBefnVZXNIeuJqNoZWDGTRGNU";
const postalCardNumber = "0079158269";
const authURL = "https://api.correios.com.br/token/v1/autentica/cartaopostagem";
const base64 = Buffer.from(`${username}:${password}`).toString("base64");

// Função para autenticar e obter token
async function autenticarCorreios() {
  const response = await fetch(authURL, {
    method: "POST",
    headers: {
      "Authorization": `Basic ${base64}`,
      "Content-Type": "application/json",
      "Accept": "application/json",
    },
    body: JSON.stringify({ numero: postalCardNumber }),
  });

  const texto = await response.text();
  if (response.status === 201) {
    const data = JSON.parse(texto);
    return `Bearer ${data.token}`;
  } else {
    throw new Error("Erro ao autenticar nos Correios");
  }
}

// Função para buscar rastreios
async function buscarPPNDetalhes(codigos, token) {
  const resultados = [];

  for (const codigo of codigos) {
    try {
      const url = `https://api.correios.com.br/prepostagem/v2/prepostagens?codigoObjeto=${codigo}`;
      const response = await fetch(url, {
        method: "GET",
        headers: {
          "Authorization": token,
          "Content-Type": "application/json",
        },
      });

      const json = await response.json();

      if (response.ok) {
        resultados.push({ codigo, success: true, data: json });
      } else {
        resultados.push({
          codigo,
          success: false,
          error: json.messages || json.message || JSON.stringify(json),
        });
      }
    } catch (err) {
      resultados.push({ codigo, success: false, error: err.message });
    }
  }

  return resultados;
}

// Rota GET /rastrear?codigos=AA123456789BR,BB987654321BR
app.get('/rastrear', async (req, res) => {
  const { codigos } = req.query;

  if (!codigos) {
    return res.status(400).json({ error: "Parâmetro 'codigos' é obrigatório. Ex: /rastrear?codigos=XX123BR" });
  }

  const listaCodigos = codigos.split(',').map(c => c.trim());

  try {
    const token = await autenticarCorreios();
    const resultados = await buscarPPNDetalhes(listaCodigos, token);
    res.json(resultados);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// NOVO endpoint para baixar a planilha XLSX dos rastreios
// Rota GET /baixar-planilha?codigos=AA123456789BR,BB987654321BR
app.get('/baixar-planilha', async (req, res) => {
  const { codigos } = req.query;

  if (!codigos) {
    return res.status(400).json({ error: "Parâmetro 'codigos' é obrigatório. Ex: /baixar-planilha?codigos=XX123BR" });
  }

  const listaCodigos = codigos.split(',').map(c => c.trim());

  try {
    const token = await autenticarCorreios();
    const resultados = await buscarPPNDetalhes(listaCodigos, token);

    // Transformar os dados para um formato simples para a planilha
    const planilhaDados = resultados.map(item => {
      if (item.success) {
        // Você pode ajustar aqui conforme o que deseja extrair do JSON de dados
        return {
          Código: item.codigo,
          Status: JSON.stringify(item.data.status || item.data), // Exemplo simplificado
          // Pode extrair outras propriedades aqui
        };
      } else {
        return {
          Código: item.codigo,
          Status: "Erro",
          MensagemErro: item.error,
        };
      }
    });

    // Criar worksheet e workbook
    const ws = xlsx.utils.json_to_sheet(planilhaDados);
    const wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, ws, 'Rastreios');

    // Gerar buffer XLSX
    const buffer = xlsx.write(wb, { type: 'buffer', bookType: 'xlsx' });

    // Cabeçalhos para forçar download do arquivo
    res.setHeader('Content-Disposition', 'attachment; filename=rastreios.xlsx');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

    // Enviar o arquivo para o cliente
    res.send(buffer);

  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Inicia servidor
app.listen(PORT, () => {
  console.log(`🚀 Servidor rodando na porta ${PORT}`);
});
