import express from "express";
import cors from "cors";
import fetch from "node-fetch";
import xlsx from "xlsx";

const app = express();
app.use(cors());
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
      Authorization: `Basic ${base64}`,
      "Content-Type": "application/json",
      Accept: "application/json",
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
          Authorization: token,
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

// Rota GET /rastrear?codigos=...
app.get("/rastrear", async (req, res) => {
  const { codigos } = req.query;

  if (!codigos) {
    return res.status(400).json({
      error: "Parâmetro 'codigos' é obrigatório. Ex: /rastrear?codigos=XX123BR",
    });
  }

  const listaCodigos = codigos.split(",").map((c) => c.trim());

  try {
    const token = await autenticarCorreios();
    const resultados = await buscarPPNDetalhes(listaCodigos, token);
    res.json(resultados);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Rota GET /baixar-planilha?codigos=...
app.get("/baixar-planilha", async (req, res) => {
  const { codigos } = req.query;

  if (!codigos) {
    return res.status(400).json({
      error:
        "Parâmetro 'codigos' é obrigatório. Ex: /baixar-planilha?codigos=XX123BR",
    });
  }

  const listaCodigos = codigos.split(",").map((c) => c.trim());

  try {
    const token = await autenticarCorreios();
    const resultados = await buscarPPNDetalhes(listaCodigos, token);

    resultados.forEach((item) => {
      if (
        item.success &&
        item.data &&
        item.data.itens &&
        item.data.itens.length > 0
      ) {
        console.log("Dados do objeto:", item.data.itens[0]);
      } else {
        console.log(
          `Erro ou dados não encontrados para o código: ${item.codigo}`
        );
      }
    });

    // Transformar os dados para a planilha com base em item.data.itens[0]
    const planilhaDados = resultados.map((item) => {
      if (item.success) {
        const dados = item.data?.itens?.[0];

        if (!dados) {
          return {
            "Código do Objeto": item.codigo,
            Status: "Não encontrado",
            Remetente: "---",
            Destinatário: "---",
            UF: "---",
            Cidade: "---",
            "Data do Status": "---",
            Serviço: "---",
            "Peso (g)": "---",
          };
        }

        return {
          "Código do Objeto": dados.codigoObjeto || item.codigo,
          Status: dados.descStatusAtual || "---",
          Remetente: dados.remetente?.nome || "---",
          "Remetente CNPJ": dados.remetente?.cpfCnpj || "---",
          "Logradouro Remetente":
            dados.remetente?.endereco?.logradouro || "---",
          "Número Remetente": dados.remetente?.endereco?.numero || "---",
          "Complemento Remetente":
            dados.remetente?.endereco?.complemento || "---",
          "Bairro Remetente": dados.remetente?.endereco?.bairro || "---",
          "Cidade Remetente": dados.remetente?.endereco?.cidade || "---",
          "CEP Remetente": dados.remetente?.endereco?.cep || "---",
          "UF Remetente": dados.remetente?.endereco?.uf || "---",
          Destinatário: dados.destinatario?.nome || "---",
          "UF Destinatário": dados.destinatario?.endereco?.uf || "---",
          "Cidade Destinatário": dados.destinatario?.endereco?.cidade || "---",
          "Data do Status": dados.dataHoraStatusAtual?.split("T")[0] || "---",
          Serviço: dados.servico || "---",
          "Peso Informado (g)": dados.pesoInformado || "---",
          "Peso Aferição (g)": dados.pesoPreAfericao || "---",
          "Altura Informada": dados.alturaInformada || "---",
          "Altura Aferição": dados.alturaPreAfericao || "---",
          "Largura Informada": dados.larguraInformada || "---",
          "Largura Aferição": dados.larguraPreAfericao || "---",
          "Comprimento Informado": dados.comprimentoInformado || "---",
          "Comprimento Aferição": dados.comprimentoPreAfericao || "---",
        };
      } else {
        return {
          "Código do Objeto": dados.codigoObjeto || item.codigo,
          Status: dados.descStatusAtual || "---",
          Remetente: dados.remetente?.nome || "---",
          "Remetente CNPJ": dados.remetente?.cpfCnpj || "---",
          "Logradouro Remetente":
            dados.remetente?.endereco?.logradouro || "---",
          "Número Remetente": dados.remetente?.endereco?.numero || "---",
          "Complemento Remetente":
            dados.remetente?.endereco?.complemento || "---",
          "Bairro Remetente": dados.remetente?.endereco?.bairro || "---",
          "Cidade Remetente": dados.remetente?.endereco?.cidade || "---",
          "CEP Remetente": dados.remetente?.endereco?.cep || "---",
          "UF Remetente": dados.remetente?.endereco?.uf || "---",
          Destinatário: dados.destinatario?.nome || "---",
          "UF Destinatário": dados.destinatario?.endereco?.uf || "---",
          "Cidade Destinatário": dados.destinatario?.endereco?.cidade || "---",
          "Data do Status": dados.dataHoraStatusAtual?.split("T")[0] || "---",
          Serviço: dados.servico || "---",
          "Peso Informado (g)": dados.pesoInformado || "---",
          "Peso Aferição (g)": dados.pesoPreAfericao || "---",
          "Altura Informada": dados.alturaInformada || "---",
          "Altura Aferição": dados.alturaPreAfericao || "---",
          "Largura Informada": dados.larguraInformada || "---",
          "Largura Aferição": dados.larguraPreAfericao || "---",
          "Comprimento Informado": dados.comprimentoInformado || "---",
          "Comprimento Aferição": dados.comprimentoPreAfericao || "---",
        };
      }
    });

    // Criar worksheet e workbook
    const ws = xlsx.utils.json_to_sheet(planilhaDados);
    const wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, ws, "Rastreios");

    // Gerar buffer XLSX
    const buffer = xlsx.write(wb, { type: "buffer", bookType: "xlsx" });

    // Cabeçalhos para forçar download
    res.setHeader("Content-Disposition", "attachment; filename=rastreios.xlsx");
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.send(buffer);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Função para buscar status detalhado de objetos (SRORastro)
async function buscarStatusObjetos(codigos, token) {
  const resultados = [];

  for (const codigo of codigos) {
    try {
      const url = `https://api.correios.com.br/srorastro/v1/objetos/${codigo}?resultado=T`;

      const response = await fetch(url, {
        method: "GET",
        headers: {
          Authorization: token,
          Accept: "application/json",
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

// Rota GET /status?codigos=...
app.get("/status", async (req, res) => {
  const { codigos } = req.query;

  if (!codigos) {
    return res.status(400).json({
      error: "Parâmetro 'codigos' é obrigatório. Ex: /status?codigos=XX123BR",
    });
  }

  const listaCodigos = codigos.split(",").map((c) => c.trim());

  try {
    const token = await autenticarCorreios();
    const resultados = await buscarStatusObjetos(listaCodigos, token);

    const respostaFormatada = resultados.map((item) => {
      if (
        item.success &&
        item.data?.objetos &&
        item.data.objetos.length > 0 &&
        item.data.objetos[0].eventos &&
        item.data.objetos[0].eventos.length > 0
      ) {
        const eventos = item.data.objetos[0].eventos;
        const ultimoEvento = eventos[0]; // geralmente o mais recente

        return {
          codigo: item.codigo,
          status: ultimoEvento.descricao || "---",
          data: ultimoEvento.dtHrCriado || "---",
          local: ultimoEvento.unidade?.endereco?.cidade || "---",
          uf: ultimoEvento.unidade?.endereco?.uf || "---",
        };
      } else {
        return {
          codigo: item.codigo,
          status: "Não encontrado ou sem eventos",
        };
      }
    });

    res.json(respostaFormatada);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Rota GET /baixar-planilha-status?codigos=...
app.get("/baixar-planilha-status", async (req, res) => {
  const { codigos } = req.query;

  if (!codigos) {
    return res.status(400).json({
      error:
        "Parâmetro 'codigos' é obrigatório. Ex: /baixar-planilha-status?codigos=XX123BR",
    });
  }

  const listaCodigos = codigos.split(",").map((c) => c.trim());

  try {
    const token = await autenticarCorreios();
    const resultados = await buscarStatusObjetos(listaCodigos, token);

    const planilhaDados = resultados.map((item) => {
      if (
        item.success &&
        item.data?.objetos &&
        item.data.objetos.length > 0 &&
        item.data.objetos[0].eventos &&
        item.data.objetos[0].eventos.length > 0
      ) {
        const eventos = item.data.objetos[0].eventos;
        const ultimoEvento = eventos[0];

        return {
          "Código do Objeto": item.codigo,
          Status: ultimoEvento.descricao || "---",
          "Data do Status": ultimoEvento.dtHrCriado || "---",
          Local: ultimoEvento.unidade?.endereco?.cidade || "---",
          UF: ultimoEvento.unidade?.endereco?.uf || "---",
        };
      } else {
        return {
          "Código do Objeto": item.codigo,
          Status: "Não encontrado ou sem eventos",
          "Data do Status": "---",
          Local: "---",
          UF: "---",
        };
      }
    });

    // Criar worksheet e workbook
    const ws = xlsx.utils.json_to_sheet(planilhaDados);
    const wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, ws, "StatusRastreio");

    // Gerar buffer XLSX
    const buffer = xlsx.write(wb, { type: "buffer", bookType: "xlsx" });

    // Cabeçalhos para forçar download
    res.setHeader(
      "Content-Disposition",
      "attachment; filename=status-rastreio.xlsx"
    );
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.send(buffer);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Inicia servidor
app.listen(PORT, () => {
  console.log(`🚀 Servidor rodando na porta ${PORT}`);
});
