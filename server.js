const express = require("express");
const { chromium } = require("playwright");
const fs = require("fs");
const path = require("path");
const { fork } = require("child_process");
const ExcelJS = require("exceljs");
const sharp = require("sharp");

const app = express();
const PORT = 3000;

const URL =
    "https://simam.tce.pr.gov.br/Paginas/Rel_LRF.aspx?relTipo=1";

app.use(express.json());
app.use(express.static(path.join(__dirname, "PUBLIC")));

app.get("/", (req, res) => {
    res.sendFile(path.join(__dirname, "PUBLIC", "index.html"));
});

// ======================================================
// PLAYWRIGHT — controle do navegador e estado da sessão
// ======================================================

let browser = null;
let context = null;
let page = null;

let municipiosCache = [];
let pronto = false;

let consultaEmAndamento = false;

let filaConsultas = Promise.resolve();
let limpezaHistoricoPendente = Promise.resolve();
let workersHistoricoAtivos = 0;
let aguardandoWorkersHistorico = [];

let configuracoesCache = new Map();
let fundebConfiguracoesCache = new Map();
let historicoExportCache = new Map();

// ======================================================
// CONFIGURAÇÕES — parâmetros gerais
// ======================================================

const ANO_INICIAL = 2019;

// ======================================================
// ANO / MÊS — período atual
// ======================================================

function obterAnoAtual() {
    return new Date().getFullYear();
}

function obterMesAtual() {
    return new Date().getMonth() + 1;
}

// ======================================================
// GERAR ANOS — anos disponíveis no sistema
// ======================================================

function gerarAnos() {
    const anoAtual = Math.min(obterAnoAtual(), 2025);
    const anos = [];

    for (let ano = ANO_INICIAL; ano <= anoAtual; ano++) {
        anos.push(String(ano));
    }

    return anos;
}

// ======================================================
// NÚMERO BRASILEIRO — conversão de valores
// ======================================================

function numeroBR(valor) {
    if (valor === null || valor === undefined) {
        return NaN;
    }

    let texto = String(valor)
        .replace(/\s/g, "")
        .trim();

    if (!texto) {
        return NaN;
    }

    texto = texto.replace(/[R$]/g, "");

    if (texto.includes(",")) {
        texto = texto
            .replace(/\./g, "")
            .replace(",", ".");
    } else {
        texto = texto.replace(/[^\d-]/g, "");
    }

    return Number(texto);
}

// ======================================================
// FORMATAR BR — apresentação de valores
// ======================================================

function formatarBR(valor) {
    if (
        valor === null ||
        valor === undefined ||
        Number.isNaN(Number(valor))
    ) {
        return "-";
    }

    return Number(valor).toLocaleString("pt-BR", {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    });
}

// ======================================================
// LIMPAR TEXTO — limpeza do conteúdo recebido
// ======================================================

function limparTexto(texto) {
    return String(texto)
        .replace(/\u00A0/g, " ")
        .replace(/\s+/g, " ")
        .trim();
}

// ======================================================
// NORMALIZAR TEXTO — comparação padronizada
// ======================================================

function normalizarTexto(texto) {
    return String(texto)
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .toUpperCase()
        .replace(/\s+/g, " ")
        .trim();
}

// ======================================================
// ENCONTRAR NÚMERO BRASILEIRO — localização de valores
// ======================================================

function encontrarNumeroBR(texto) {
    if (!texto) {
        return null;
    }

    const encontrados = String(texto).match(
        /-?\d{1,3}(?:\.\d{3})*,\d{2}/g
    );

    if (!encontrados || encontrados.length === 0) {
        return null;
    }

    return encontrados[0];
}

// ======================================================
// EXTRAIR VALOR POR MARCADOR — busca por marcador
// ======================================================

function extrairValorPorMarcador(texto, marcadores) {
    const textoNormalizado = limparTexto(texto);

    for (const marcador of marcadores) {
        const indice = textoNormalizado
            .toUpperCase()
            .indexOf(marcador.toUpperCase());

        if (indice === -1) {
            continue;
        }

        const trecho = textoNormalizado.substring(
            indice + marcador.length
        );

        const numero = encontrarNumeroBR(trecho);

        if (numero) {
            return numero;
        }
    }

    return null;
}

// ======================================================
// EXTRAIR FUNDEB — leitura do relatório MDE
// ======================================================

// Mantém o valor realizado do FUNDEB conforme o relatório do TCE.
function extrairFundeb(texto) {
    if (!texto) {
        return null;
    }

    console.log("");
    console.log("=================================");
    console.log("EXTRAINDO TRANSFERÊNCIAS DO FUNDEB");
    console.log("=================================");

    const textoOriginal = String(texto)
        .replace(/^\uFEFF/, "")
        .replace(/\r/g, "")
        .replace(/\u00A0/g, " ");

    // ======================================================
// FORMATO DO CSV DO TCE-PR — estrutura esperada
    // ======================================================
// EXEMPLO DO CSV — referência de estrutura
    // ======================================================

    const regexNumeroBR = /-?(?:\d{1,3}(?:\.\d{3})+|\d+),\d{2}/g;

    function dividirCSV(linha) {
        const resultado = [];
        let campo = "";
        let dentroDeAspas = false;

        for (let i = 0; i < linha.length; i++) {
            const caractere = linha[i];

            if (caractere === '"') {
                if (
                    dentroDeAspas &&
                    linha[i + 1] === '"'
                ) {
                    campo += '"';
                    i++;
                    continue;
                }

                dentroDeAspas = !dentroDeAspas;
                continue;
            }

            const temPontoEVirgula = linha.includes(";");
            const temTabulacao = linha.includes("\t");
            const separador = temPontoEVirgula
                ? ";"
                : temTabulacao
                    ? "\t"
                    : ",";

            if (
                caractere === separador &&
                !dentroDeAspas
            ) {
                resultado.push(campo.trim());
                campo = "";
                continue;
            }

            campo += caractere;
        }

        resultado.push(campo.trim());
        return resultado;
    }

    function extrairNumeros(linha) {
        if (!linha) {
            return [];
        }

        return String(linha).match(regexNumeroBR) || [];
    }

    function retornarFundeb(valorTexto, origem) {
        if (!valorTexto) {
            return null;
        }

        const valor = numeroBR(valorTexto);

        if (Number.isNaN(valor)) {
            return null;
        }

        console.log(
            "Transferências do FUNDEB:",
            formatarBR(valor)
        );

        console.log("Origem:", origem);
        console.log("=================================");

        return {
            fundeb: valor,
            fundebFormatado: formatarBR(valor),
            fundebOrigem: origem
        };
    }

    // ======================================================
// ESCOLHER O REALIZADO — identifica o valor realizado
    // ======================================================
    function escolherValorRealizado(linha, colunas) {
        const numeros = extrairNumeros(linha);

        if (numeros.length === 0) {
            return null;
        }

        const usaSeparadorEstrutural =
            String(linha).includes(";") ||
            String(linha).includes("\t");

        if (usaSeparadorEstrutural && colunas && colunas.length >= 4) {
            const candidatos = [
                colunas[3],
                colunas[2]
            ];

            for (const candidato of candidatos) {
                const valor = numeroBR(candidato);

                if (!Number.isNaN(valor)) {
                    return candidato;
                }
            }
        }

        if (numeros.length >= 4) {
            return numeros[2];
        }

        if (numeros.length === 3) {
            return numeros[1];
        }

        if (numeros.length === 2) {
            return numeros[1];
        }

        return numeros[0];
    }

    const linhas = textoOriginal
        .split("\n")
        .map(linha => linha.trim())
        .filter(Boolean);

    // ======================================================
// FUNDEB — identificação da receita recebida
    // ======================================================
// FUNDEB — aceita variações de numeração
    // ======================================================
    const marcadoresPrincipais = [
        "RECEITAS RECEBIDAS DO FUNDEB",
        "RECEITAS DO FUNDEB",
        "RECEITAS DESTINADAS AO FUNDEB",
        "TOTAL DAS RECEITAS TRANSFERIDAS AO FUNDEB",
        "TOTAL DESTINADO AO FUNDEB"
    ];

    for (let i = 0; i < linhas.length; i++) {
        const linha = linhas[i];
        const normalizada = normalizarTexto(linha);

        const marcador = marcadoresPrincipais.find(item =>
            normalizada.includes(normalizarTexto(item))
        );

        if (!marcador) {
            continue;
        }

        const colunas = dividirCSV(linha);
        const numeros = extrairNumeros(linha);

        if (numeros.length >= 2) {
            const numero = escolherValorRealizado(linha, colunas);

            console.log("");
            console.log("Linha do FUNDEB encontrada:");
            console.log(linha);
            console.log("Colunas detectadas:", colunas);
            console.log("Números encontrados:", numeros);
            console.log("Valor realizado escolhido:", numero || "-");

            const resultado = retornarFundeb(
                numero,
                `${linha.split(",")[0]} - valor realizado`
            );

            if (resultado) {
                return resultado;
            }
        }

        const blocoLinhas = [linha];

        for (
            let j = i + 1;
            j < linhas.length && blocoLinhas.length < 8;
            j++
        ) {
            const proxima = linhas[j];
            const normalizadaProxima = normalizarTexto(proxima);

            if (
                /^\d+(?:\.\d+)*\s*-\s*/.test(normalizadaProxima)
            ) {
                break;
            }

            blocoLinhas.push(proxima);

            const numerosBloco = extrairNumeros(
                blocoLinhas.join(" | ")
            );

            if (numerosBloco.length >= 4 || /%/.test(proxima)) {
                break;
            }
        }

        const bloco = blocoLinhas.join(" | ");
        const numerosBloco = extrairNumeros(bloco);

        if (numerosBloco.length === 0) {
            continue;
        }

        let numero = null;

        if (numerosBloco.length >= 4) {
            numero = numerosBloco[2];
        } else if (numerosBloco.length === 3) {
            numero = numerosBloco[1];
        } else if (numerosBloco.length === 2) {
            numero = numerosBloco[1];
        } else {
            numero = numerosBloco[0];
        }

        console.log("");
        console.log("Bloco FUNDEB encontrado:");
        console.log(bloco);
        console.log("Números encontrados:", numerosBloco);
        console.log("Valor realizado escolhido:", numero || "-");

        const resultado = retornarFundeb(
            numero,
            `${marcador} - valor realizado`
        );

        if (resultado) {
            return resultado;
        }
    }

    // ======================================================
// FUNDEB — fallback para texto ou HTML
    // ======================================================
// FUNDEB — fallback quando o relatório chega como HTML
    // ======================================================
    const textoLimpo = textoOriginal
        .replace(/<script[\s\S]*?<\/script>/gi, " ")
        .replace(/<style[\s\S]*?<\/style>/gi, " ")
        .replace(/<br\s*\/?>/gi, "\n")
        .replace(/<\/(?:td|th|tr|p|div|li)>/gi, "\n")
        .replace(/<[^>]+>/g, " ")
        .replace(/&nbsp;|&#160;|&#xA0;/gi, " ")
        .replace(/&amp;/gi, "&")
        .replace(/&quot;/gi, '"')
        .replace(/&#39;|&apos;/gi, "'")
        .replace(/&lt;/gi, "<")
        .replace(/&gt;/gi, ">");

    const linhasHTML = textoLimpo
        .split("\n")
        .map(linha => linha.replace(/[ \t]+/g, " ").trim())
        .filter(Boolean);

    for (let i = 0; i < linhasHTML.length; i++) {
        const normalizada = normalizarTexto(linhasHTML[i]);

        if (
            !normalizada.includes("RECEITAS RECEBIDAS DO FUNDEB") &&
            !normalizada.includes("RECEITAS DO FUNDEB") &&
            !normalizada.includes("RECEITAS DESTINADAS AO FUNDEB") &&
            !normalizada.includes("TOTAL DAS RECEITAS TRANSFERIDAS AO FUNDEB") &&
            !normalizada.includes("TOTAL DESTINADO AO FUNDEB")
        ) {
            continue;
        }

        const blocoLinhas = [linhasHTML[i]];

        for (
            let j = i + 1;
            j < linhasHTML.length && blocoLinhas.length < 8;
            j++
        ) {
            const proxima = linhasHTML[j];
            const normalizadaProxima = normalizarTexto(proxima);

            if (
                /^\d+(?:\.\d+)*\s*-\s*/.test(normalizadaProxima)
            ) {
                break;
            }

            blocoLinhas.push(proxima);

            const nums = extrairNumeros(
                blocoLinhas.join(" | ")
            );

            if (nums.length >= 4 || /%/.test(proxima)) {
                break;
            }
        }

        const bloco = blocoLinhas.join(" | ");
        const numeros = extrairNumeros(bloco);

        if (numeros.length === 0) {
            continue;
        }

        let numero;

        if (numeros.length >= 4) {
            numero = numeros[2];
        } else if (numeros.length === 3) {
            numero = numeros[1];
        } else if (numeros.length === 2) {
            numero = numeros[1];
        } else {
            numero = numeros[0];
        }

        console.log("");
        console.log("Bloco FUNDEB HTML encontrado:");
        console.log(bloco);
        console.log("Números encontrados:", numeros);
        console.log("Valor realizado escolhido:", numero || "-");

        const resultado = retornarFundeb(
            numero,
            "FUNDEB - valor realizado extraído do relatório MDE"
        );

        if (resultado) {
            return resultado;
        }
    }

    // ======================================================
// FUNDEB — fallback para variações recentes
    // ======================================================
// FUNDEB — compatibilidade com variações do relatório
    const marcadoresFlexiveis = [
        "TRANSFERENCIAS RECEBIDAS DO FUNDEB",
        "TRANSFERENCIAS DO FUNDEB",
        "RECEITAS RECEBIDAS FUNDEB",
        "RECEITAS FUNDEB",
        "VALOR RECEBIDO DO FUNDEB",
        "VALOR RECEBIDO FUNDEB",
        "RECEITA DO FUNDEB"
    ];

    const linhasFallback = [
        ...linhas,
        ...linhasHTML
    ];

    for (let i = 0; i < linhasFallback.length; i++) {
        const normalizada = normalizarTexto(linhasFallback[i]);

        if (!normalizada.includes("FUNDEB")) {
            continue;
        }

        const temMarcadorFlexivel = marcadoresFlexiveis.some(
            marcador => normalizada.includes(normalizarTexto(marcador))
        );

        if (!temMarcadorFlexivel) {
            continue;
        }

        const blocoLinhas = [linhasFallback[i]];

        for (
            let j = i + 1;
            j < linhasFallback.length && blocoLinhas.length < 8;
            j++
        ) {
            const proxima = linhasFallback[j];
            const normalizadaProxima = normalizarTexto(proxima);

            if (
                /^\d+(?:\.\d+)*\s*-\s*/.test(normalizadaProxima)
            ) {
                break;
            }

            blocoLinhas.push(proxima);

            const numerosBloco = extrairNumeros(
                blocoLinhas.join(" | ")
            );

            if (numerosBloco.length >= 4 || /%/.test(proxima)) {
                break;
            }
        }

        const bloco = blocoLinhas.join(" | ");
        const numeros = extrairNumeros(bloco);

        if (numeros.length === 0) {
            continue;
        }

        let numero = null;

        if (numeros.length >= 4) {
            numero = numeros[2];
        } else if (numeros.length === 3) {
            numero = numeros[1];
        } else if (numeros.length === 2) {
            numero = numeros[1];
        } else {
            numero = numeros[0];
        }

        const resultado = retornarFundeb(
            numero,
            "FUNDEB - fallback por descrição alternativa"
        );

        if (resultado) {
            return resultado;
        }
    }

    console.log(
        "Valor de Transferências do FUNDEB não encontrado."
    );
    console.log("=================================");

    return null;
}


// ======================================================
// EXTRAIR DTP — tabela do limite legal
// ======================================================

function extrairDtpTabelaLimite(texto) {
    if (!texto) {
        return null;
    }

    const textoNormalizado = normalizarTexto(texto);

    const marcadorTabela =
        "APURACAO DO CUMPRIMENTO DO LIMITE LEGAL";

    const marcadorDtp =
        "DESPESA TOTAL COM PESSOAL - DTP";

    const regexNumero =
        /-?(?:\d{1,3}(?:\.\d{3})+|\d+),\d{2}/g;

    const inicioTabela =
        textoNormalizado.indexOf(marcadorTabela);

    if (inicioTabela !== -1) {
        const trechoTabela =
            textoNormalizado.substring(inicioTabela);

        const inicioDtp =
            trechoTabela.indexOf(marcadorDtp);

        if (inicioDtp !== -1) {
            const trechoDtp =
                trechoTabela.substring(
                    inicioDtp + marcadorDtp.length
                );

            const numeros =
                trechoDtp.match(regexNumero) || [];

            if (numeros.length > 0) {
                console.log(
                    "DTP encontrado na tabela de Apuração do Limite Legal:",
                    numeros[0]
                );

                return numeros[0];
            }
        }
    }

    let ultimaOcorrencia = -1;
    let posicaoBusca = 0;

    while (true) {
        const inicioDtp =
            textoNormalizado.indexOf(
                marcadorDtp,
                posicaoBusca
            );

        if (inicioDtp === -1) {
            break;
        }

        ultimaOcorrencia = inicioDtp;
        posicaoBusca =
            inicioDtp + marcadorDtp.length;
    }

    if (ultimaOcorrencia !== -1) {
        const trechoDepoisDtp =
            textoNormalizado.substring(
                ultimaOcorrencia + marcadorDtp.length
            );

        const numeros =
            trechoDepoisDtp.match(regexNumero) || [];

        if (numeros.length > 0) {
            console.log(
                "DTP encontrado pela última ocorrência da tabela:",
                numeros[0]
            );

            return numeros[0];
        }
    }

    console.log(
        "DTP não encontrado na tabela de Apuração do Limite Legal."
    );

    return null;
}

// ======================================================
// EXTRAIR LRF — RCL, RCL ajustada, DTP e percentual
// ======================================================

function extrairDadosRelatorio(texto) {
    const textoLimpo = limparTexto(texto);

    console.log("");
    console.log("=================================");
    console.log("EXTRAINDO DADOS DO RELATÓRIO");
    console.log("=================================");

    const rclTexto = extrairValorPorMarcador(textoLimpo, [
        "RECEITA CORRENTE LÍQUIDA - RCL (IV)",
        "RECEITA CORRENTE LÍQUIDA - RCL",
        "RECEITA CORRENTE LÍQUIDA"
    ]);

    const rclAjustadaTexto = extrairValorPorMarcador(textoLimpo, [
        "RECEITA CORRENTE LÍQUIDA AJUSTADA PARA CÁLCULO DOS LIMITES DA DESPESA COM PESSOAL",
        "RECEITA CORRENTE LÍQUIDA AJUSTADA"
    ]);

    const dtpTexto = extrairDtpTabelaLimite(texto);

    console.log("RCL:", rclTexto);
    console.log("RCL Ajustada:", rclAjustadaTexto);
    console.log("DTP:", dtpTexto);

    if (!rclTexto) {
        throw new Error(
            "Não foi possível localizar a Receita Corrente Líquida (RCL) no relatório."
        );
    }

    if (!rclAjustadaTexto) {
        throw new Error(
            "Não foi possível localizar a Receita Corrente Líquida Ajustada no relatório."
        );
    }

    if (!dtpTexto) {
        throw new Error(
            "Não foi possível localizar a Despesa Total com Pessoal (DTP) no relatório."
        );
    }

    const rcl = numeroBR(rclTexto);
    const rclAjustada = numeroBR(rclAjustadaTexto);
    const dtp = numeroBR(dtpTexto);

    if (
        Number.isNaN(rcl) ||
        Number.isNaN(rclAjustada) ||
        Number.isNaN(dtp)
    ) {
        throw new Error(
            "Erro ao converter os valores encontrados no relatório."
        );
    }

    const percentual =
        rclAjustada > 0
            ? (dtp / rclAjustada) * 100
            : 0;

    const limiteAlertaPercentual = 48.6;
    const limitePrudencialPercentual = 51.3;
    const limiteMaximoPercentual = 54;

    const limiteAlerta =
        rclAjustada *
        (limiteAlertaPercentual / 100);

    const limitePrudencial =
        rclAjustada *
        (limitePrudencialPercentual / 100);

    const limiteMaximo =
        rclAjustada *
        (limiteMaximoPercentual / 100);

    let situacao = "OK";

    if (percentual >= limiteMaximoPercentual) {
        situacao = "ACIMA DO LIMITE MÁXIMO";
    } else if (percentual >= limitePrudencialPercentual) {
        situacao = "LIMITE PRUDENCIAL";
    } else if (percentual >= limiteAlertaPercentual) {
        situacao = "LIMITE DE ALERTA";
    }

    return {
        rcl,
        rclFormatada: formatarBR(rcl),

        rclAjustada,
        rclAjustadaFormatada: formatarBR(rclAjustada),

        dtp,
        dtpFormatado: formatarBR(dtp),

        percentual,
        percentualFormatado:
            percentual.toFixed(2) + "%",

        limiteAlerta,
        limiteAlertaFormatado:
            formatarBR(limiteAlerta),
        limiteAlertaPercentual,

        limitePrudencial,
        limitePrudencialFormatado:
            formatarBR(limitePrudencial),
        limitePrudencialPercentual,

        limiteMaximo,
        limiteMaximoFormatado:
            formatarBR(limiteMaximo),
        limiteMaximoPercentual,

        situacao
    };
}

// ======================================================
// EVOLUÇÃO — variações anuais e acumuladas
// ======================================================

function calcularEvolucao(valorAtual, valorAnterior) {
    if (
        valorAtual === null ||
        valorAtual === undefined ||
        Number.isNaN(Number(valorAtual))
    ) {
        return null;
    }

    if (
        valorAnterior === null ||
        valorAnterior === undefined ||
        Number.isNaN(Number(valorAnterior)) ||
        Number(valorAnterior) === 0
    ) {
        return null;
    }

    return (
        (
            (Number(valorAtual) - Number(valorAnterior)) /
            Number(valorAnterior)
        ) * 100
    );
}

function calcularEvolucaoAcumulada(valorAtual, valorBase) {
    if (
        valorAtual === null ||
        valorAtual === undefined ||
        Number.isNaN(Number(valorAtual))
    ) {
        return null;
    }

    if (
        valorBase === null ||
        valorBase === undefined ||
        Number.isNaN(Number(valorBase)) ||
        Number(valorBase) === 0
    ) {
        return null;
    }

    return (
        (
            (Number(valorAtual) - Number(valorBase)) /
            Number(valorBase)
        ) * 100
    );
}

function formatarPercentualEvolucao(valor) {
    if (
        valor === null ||
        valor === undefined ||
        Number.isNaN(Number(valor))
    ) {
        return "-";
    }

    const numero = Number(valor);
    const sinal = numero > 0 ? "+" : "";

    return (
        sinal +
        numero.toFixed(2).replace(".", ",") +
        "%"
    );
}

// ======================================================
// HISTÓRICO — dados de 2019 a 2025
// ======================================================

function calcularHistorico(historico) {
    if (!historico || historico.length === 0) {
        return [];
    }

    const base =
        historico.find(item => Number(item.ano) === ANO_INICIAL) ||
        historico[0];

    return historico.map((item, index) => {
        const anterior =
            index > 0
                ? historico[index - 1]
                : null;

        const evolucaoRCL =
            calcularEvolucao(
                item.rcl,
                anterior ? anterior.rcl : null
            );

        const acumuladoRCL =
            calcularEvolucaoAcumulada(
                item.rcl,
                base.rcl
            );

        const evolucaoRCLAjustada =
            calcularEvolucao(
                item.rclAjustada,
                anterior
                    ? anterior.rclAjustada
                    : null
            );

        const acumuladoRCLAjustada =
            calcularEvolucaoAcumulada(
                item.rclAjustada,
                base.rclAjustada
            );

        const evolucaoDTP =
            calcularEvolucao(
                item.dtp,
                anterior ? anterior.dtp : null
            );

        const acumuladoDTP =
            calcularEvolucaoAcumulada(
                item.dtp,
                base.dtp
            );

        const evolucaoFundeb =
            calcularEvolucao(
                item.fundeb,
                anterior ? anterior.fundeb : null
            );

        const acumuladoFundeb =
            calcularEvolucaoAcumulada(
                item.fundeb,
                base.fundeb
            );

        return {
            ...item,

            evolucaoRCL,
            evolucaoRCLFormatada:
                formatarPercentualEvolucao(
                    evolucaoRCL
                ),

            acumuladoRCL,
            acumuladoRCLFormatado:
                formatarPercentualEvolucao(
                    acumuladoRCL
                ),

            evolucaoRCLAjustada,
            evolucaoRCLAjustadaFormatada:
                formatarPercentualEvolucao(
                    evolucaoRCLAjustada
                ),

            acumuladoRCLAjustada,
            acumuladoRCLAjustadaFormatado:
                formatarPercentualEvolucao(
                    acumuladoRCLAjustada
                ),

            evolucaoDTP,
            evolucaoDTPFormatada:
                formatarPercentualEvolucao(
                    evolucaoDTP
                ),

            acumuladoDTP,
            acumuladoDTPFormatado:
                formatarPercentualEvolucao(
                    acumuladoDTP
                ),

            evolucaoFundeb,
            evolucaoFundebFormatada:
                formatarPercentualEvolucao(
                    evolucaoFundeb
                ),

            acumuladoFundeb,
            acumuladoFundebFormatado:
                formatarPercentualEvolucao(
                    acumuladoFundeb
                )
        };
    });
}

// ======================================================
// ESPERAR OPÇÃO — aguarda uma opção do TCE
// ======================================================

async function esperarOpcao(
    seletor,
    valor,
    timeout = 30000
) {
    await page.waitForFunction(
        ({ seletor, valor }) => {
            const select =
                document.querySelector(seletor);

            if (!select) {
                return false;
            }

            return Array.from(
                select.options
            ).some(
                (option) =>
                    String(option.value) ===
                    String(valor)
            );
        },
        { seletor, valor },
        { timeout }
    );
}

// ======================================================
// ESPERAR QUANTIDADE DE OPÇÕES — aguarda o seletor
// ======================================================

async function esperarQuantidadeOpcoes(
    seletor,
    quantidade = 2,
    timeout = 30000
) {
    await page.waitForFunction(
        ({ seletor, quantidade }) => {
            const select =
                document.querySelector(seletor);

            if (!select) {
                return false;
            }

            return (
                select.options.length >=
                quantidade
            );
        },
        { seletor, quantidade },
        { timeout }
    );
}

// ======================================================
// SELECIONAR OPÇÃO — seleciona um item
// ======================================================

async function selecionarOpcao(
    seletor,
    valor,
    seletorDependente = null
) {
    await esperarOpcao(
        seletor,
        String(valor)
    );

    await page
        .locator(seletor)
        .selectOption(String(valor));

    await page.waitForTimeout(100);

    if (seletorDependente) {
        await esperarQuantidadeOpcoes(
            seletorDependente,
            2,
            30000
        );
    }
}

// ======================================================
// INICIAR TCE — inicialização do navegador
// ======================================================

async function iniciarTCE() {
    console.log("");
    console.log("=================================");
    console.log("INICIANDO TCE-PR");
    console.log("=================================");

    browser = await chromium.launch({
        headless: true
    });

    context = await browser.newContext();
    page = await context.newPage();

    await page.goto(URL, {
        waitUntil: "domcontentloaded",
        timeout: 60000
    });

    await page.waitForTimeout(2000);

    const tipo =
        page.locator(
            "#ContentPlaceHolder1_ddlTipo"
        );

    await tipo.selectOption("1");

    await page.waitForTimeout(1500);

    await esperarQuantidadeOpcoes(
        "#ContentPlaceHolder1_ddlMunicipio",
        2,
        30000
    );

    municipiosCache =
        await page
            .locator(
                "#ContentPlaceHolder1_ddlMunicipio option"
            )
            .evaluateAll((options) => {
                return options
                    .map((option) => ({
                        id: option.value,
                        nome:
                            option.textContent.trim()
                    }))
                    .filter(
                        (option) =>
                            option.id &&
                            option.nome &&
                            !option.nome
                                .toUpperCase()
                                .includes("SELECIONE")
                    );
            });

    console.log(
        "Municípios encontrados:",
        municipiosCache.length
    );

    pronto = true;

    console.log("TCE-PR pronto.");
}

// ======================================================
// STATUS — estado da aplicação
// ======================================================

app.get("/api/status", (req, res) => {
    res.json({
        sucesso: true,
        pronto,
        municipios:
            municipiosCache.length,
        anoAtual:
            obterAnoAtual(),
        mesAtual:
            obterMesAtual(),
        anos:
            gerarAnos()
    });
});

// ======================================================
// MUNICÍPIOS — dados disponíveis para o frontend
// ======================================================

app.get("/api/municipios", (req, res) => {
    res.json({
        sucesso: true,
        municipios:
            municipiosCache
    });
});

// ======================================================
// ENTIDADES — entidades do município
// ======================================================

app.get(
    "/api/entidades",
    async (req, res) => {
        try {
            const municipioId =
                req.query.municipio;

            if (!municipioId) {
                return res.status(400).json({
                    sucesso: false,
                    erro:
                        "Município não informado."
                });
            }

            await selecionarOpcao(
                "#ContentPlaceHolder1_ddlMunicipio",
                municipioId,
                "#ContentPlaceHolder1_ddlEntidade"
            );

            const entidades =
                await page
                    .locator(
                        "#ContentPlaceHolder1_ddlEntidade option"
                    )
                    .evaluateAll((options) => {
                        return options
                            .map((option) => ({
                                id: option.value,
                                nome:
                                    option.textContent.trim()
                            }))
                            .filter(
                                (option) =>
                                    option.id &&
                                    option.nome &&
                                    !option.nome
                                        .toUpperCase()
                                        .includes(
                                            "SELECIONE"
                                        )
                            );
                    });

            res.json({
                sucesso: true,
                entidades
            });
        } catch (erro) {
            console.error(
                "Erro ao carregar entidades:",
                erro.message
            );

            res.status(500).json({
                sucesso: false,
                erro:
                    erro.message
            });
        }
    }
);

// ======================================================
// RELATÓRIOS — relatórios usados na consulta
// ======================================================

app.get(
    "/api/relatorios",
    async (req, res) => {
        try {
            const municipioId =
                req.query.municipio;

            const entidadeId =
                req.query.entidade;

            if (!municipioId) {
                return res.status(400).json({
                    sucesso: false,
                    erro:
                        "Município não informado."
                });
            }

            if (!entidadeId) {
                return res.status(400).json({
                    sucesso: false,
                    erro:
                        "Entidade não informada."
                });
            }

            await selecionarOpcao(
                "#ContentPlaceHolder1_ddlMunicipio",
                municipioId,
                "#ContentPlaceHolder1_ddlEntidade"
            );

            await selecionarOpcao(
                "#ContentPlaceHolder1_ddlEntidade",
                entidadeId,
                "#ContentPlaceHolder1_ddlRelatorio"
            );

            const relatorios =
                await page
                    .locator(
                        "#ContentPlaceHolder1_ddlRelatorio option"
                    )
                    .evaluateAll((options) => {
                        return options
                            .map((option) => ({
                                id: option.value,
                                nome:
                                    option.textContent.trim()
                            }))
                            .filter(
                                (option) =>
                                    option.id &&
                                    option.nome &&
                                    !option.nome
                                        .toUpperCase()
                                        .includes(
                                            "SELECIONE"
                                        )
                            );
                    });

            res.json({
                sucesso: true,
                relatorios
            });
        } catch (erro) {
            console.error(
                "Erro ao carregar relatórios:",
                erro.message
            );

            res.status(500).json({
                sucesso: false,
                erro:
                    erro.message
            });
        }
    }
);

// ======================================================
// ENTIDADE PARA FUNDEB — configuração do MDE
// ======================================================

async function encontrarEntidadeFundeb() {
    const entidades =
        await page
            .locator(
                "#ContentPlaceHolder1_ddlEntidade option"
            )
            .evaluateAll((options) => {
                return options
                    .map((option) => ({
                        id: option.value,
                        nome:
                            option.textContent.trim()
                    }))
                    .filter(
                        (option) =>
                            option.id &&
                            option.nome
                    );
            });

    const candidatos =
        entidades
            .map((entidade) => {
                const nome =
                    normalizarTexto(
                        entidade.nome
                    );

                let pontos = 0;

                if (
                    nome.includes("PREFEITURA")
                ) pontos += 1000;

                if (
                    nome.startsWith("MUNICIPIO")
                ) pontos += 900;

                if (
                    nome.includes("MUNICIPIO")
                ) pontos += 500;

                if (
                    nome.includes("EXECUTIVO")
                ) pontos += 300;

                if (
                    nome.includes(
                        "FUNDO DE PREVIDENCIA"
                    )
                ) pontos -= 2000;

                if (
                    nome.includes("PREVIDENCIA")
                ) pontos -= 2000;

                if (
                    nome.includes("CAMARA")
                ) pontos -= 2000;

                if (
                    nome.includes("LEGISLATIVO")
                ) pontos -= 2000;

                if (
                    nome.includes("FUNDO")
                ) pontos -= 1000;

                if (
                    nome.includes("INSTITUTO")
                ) pontos -= 1000;

                return {
                    ...entidade,
                    pontos
                };
            })
            .sort(
                (a, b) =>
                    b.pontos - a.pontos
            );

    console.log(
        "RANKING DAS ENTIDADES PARA FUNDEB:"
    );

    candidatos.forEach((entidade) => {
        console.log(
            `${entidade.pontos} pontos -> ${entidade.nome}`
        );
    });

    const melhor =
        candidatos.find(
            (entidade) =>
                entidade.pontos > 0
        );

    if (!melhor) {
        return null;
    }

    console.log(
        "ENTIDADE SELECIONADA PARA FUNDEB:",
        melhor.nome
    );

    return melhor;
}

// ======================================================
// RELATÓRIO MDE — preparação da consulta
// ======================================================

async function encontrarRelatorioMDE() {
    const relatorios =
        await page
            .locator(
                "#ContentPlaceHolder1_ddlRelatorio option"
            )
            .evaluateAll((options) => {
                return options
                    .map((option) => ({
                        id: option.value,
                        nome:
                            option.textContent.trim()
                    }))
                    .filter(
                        (option) =>
                            option.id &&
                            option.nome
                    );
            });

    const candidatos =
        relatorios
            .map((relatorio) => {
                const nome =
                    normalizarTexto(
                        relatorio.nome
                    );

                let pontos = 0;

                if (
                    nome.includes("ANEXO 8")
                ) pontos += 1000;

                if (
                    nome.includes(
                        "DEMONSTRATIVO DAS RECEITAS E DESPESAS COM EDUCACAO"
                    )
                ) pontos += 900;

                if (
                    nome.includes(
                        "MANUTENCAO E DESENVOLVIMENTO DO ENSINO"
                    )
                ) pontos += 800;

                if (
                    nome.includes(
                        "RECEITAS E DESPESAS COM EDUCACAO"
                    )
                ) pontos += 700;

                if (
                    nome.includes("EDUCACAO")
                ) pontos += 300;

                if (
                    nome.includes("MDE")
                ) pontos += 500;

                return {
                    ...relatorio,
                    pontos
                };
            })
            .filter(
                (relatorio) =>
                    relatorio.pontos > 0
            )
            .sort(
                (a, b) =>
                    b.pontos - a.pontos
            );

    if (
        candidatos.length === 0
    ) {
        return null;
    }

    console.log(
        "RELATÓRIO MDE SELECIONADO:",
        candidatos[0].nome
    );

    return candidatos[0];
}

// ======================================================
// CONFIGURAÇÃO AUTOMÁTICA — encontra os seletores necessários
// ======================================================

async function encontrarConfiguracaoConsulta(municipioId) {
    const chave = String(municipioId);

    if (configuracoesCache.has(chave)) {
        return configuracoesCache.get(chave);
    }

    await selecionarOpcao(
        "#ContentPlaceHolder1_ddlMunicipio",
        chave,
        "#ContentPlaceHolder1_ddlEntidade"
    );

    await page.waitForFunction(
        () => {
            const select = document.querySelector(
                "#ContentPlaceHolder1_ddlEntidade"
            );

            if (!select) return false;

            return Array.from(select.options).some(option => {
                const texto = (option.textContent || "")
                    .trim()
                    .toUpperCase()
                    .normalize("NFD")
                    .replace(/[\u0300-\u036f]/g, "");

                return (
                    texto.includes("MUNICIPIO") ||
                    texto.includes("PREFEITURA") ||
                    texto.includes("EXECUTIVO")
                );
            });
        },
        { timeout: 30000 }
    );

    const entidades = await page
        .locator("#ContentPlaceHolder1_ddlEntidade option")
        .evaluateAll((options) =>
            options
                .map((option) => ({
                    id: option.value,
                    nome: (option.textContent || "").trim()
                }))
                .filter((option) => option.id && option.nome)
        );

    const entidade = entidades
        .map((item) => {
            const nome = normalizarTexto(item.nome);
            let pontos = 0;

            if (nome.startsWith("MUNICIPIO")) pontos += 10000;
            else if (nome.includes("MUNICIPIO")) pontos += 9000;
            else if (nome.includes("PREFEITURA")) pontos += 8000;

            if (nome.includes("EXECUTIVO")) pontos += 3000;

            if (nome.includes("CAMARA MUNICIPAL")) pontos -= 10000;
            else if (nome.includes("CAMARA")) pontos -= 9000;

            if (nome.includes("LEGISLATIVO")) pontos -= 3000;
            if (nome.includes("PREVIDENCIA")) pontos -= 3000;
            if (nome.includes("FUNDO")) pontos -= 2000;
            if (nome.includes("INSTITUTO")) pontos -= 2000;

            return { ...item, pontos };
        })
        .sort((a, b) => b.pontos - a.pontos)
        .find((item) => item.pontos > 0);

    if (!entidade) {
        throw new Error("Não foi possível identificar o Município para a consulta de RCL/DTP.");
    }

    await selecionarOpcao(
        "#ContentPlaceHolder1_ddlEntidade",
        entidade.id,
        "#ContentPlaceHolder1_ddlRelatorio"
    );

    const relatorios = await page
        .locator("#ContentPlaceHolder1_ddlRelatorio option")
        .evaluateAll((options) =>
            options
                .map((option) => ({
                    id: option.value,
                    nome: (option.textContent || "").trim()
                }))
                .filter((option) => option.id && option.nome)
        );

    const relatorio = relatorios
        .map((item) => {
            const nome = normalizarTexto(item.nome);
            let pontos = 0;

            if (nome === "RGF - DEMONSTRATIVO DA DESPESA COM PESSOAL") pontos += 5000;
            if (nome.includes("RGF")) pontos += 1000;
            if (nome.includes("DEMONSTRATIVO DA DESPESA COM PESSOAL")) pontos += 3000;
            if (nome.includes("DESPESA COM PESSOAL")) pontos += 1000;
            if (nome.includes("RREO")) pontos -= 1000;
            if (nome.includes("FUNDEB")) pontos -= 2000;
            if (nome.includes("DISPONIBILIDADE DE CAIXA")) pontos -= 2000;
            if (nome.includes("DIVIDA CONSOLIDADA")) pontos -= 2000;

            return { ...item, pontos };
        })
        .sort((a, b) => b.pontos - a.pontos)
        .find((item) => item.pontos > 0);

    if (!relatorio) {
        throw new Error("Não foi possível identificar o relatório de Despesa com Pessoal.");
    }

    const configuracao = {
        entidadeId: entidade.id,
        entidadeNome: entidade.nome,
        relatorioId: relatorio.id,
        relatorioNome: relatorio.nome
    };

    configuracoesCache.set(chave, configuracao);

    return configuracao;
}

async function obterConfiguracaoConsulta(municipioId) {
    return await encontrarConfiguracaoConsulta(municipioId);
}

// ======================================================
// LER DOWNLOAD COMO TEXTO — leitura do arquivo baixado
// ======================================================

async function lerDownloadComoTexto(download) {
    if (!download) {
        return null;
    }

    try {
        const caminho =
            await download.path();

        if (
            caminho &&
            fs.existsSync(caminho)
        ) {
            return fs.readFileSync(
                caminho,
                "utf8"
            );
        }
    } catch (erro) {
        console.log(
            "Não foi possível ler o caminho temporário do download."
        );
    }

    try {
        const stream =
            await download.createReadStream();

        if (!stream) {
            return null;
        }

        const partes = [];

        for await (const parte of stream) {
            partes.push(parte);
        }

        return Buffer
            .concat(partes)
            .toString("utf8");
    } catch (erro) {
        console.log(
            "Não foi possível ler o conteúdo do download."
        );

        return null;
    }
}

// ======================================================
// CAPTURAR CSV DO MDE — captura do arquivo
// ======================================================

async function capturarDownloadDoRelatorio(
    paginaRelatorio
) {
    if (!paginaRelatorio) {
        return null;
    }

    const seletorElementos =
        "a, button, input[type='button'], input[type='submit']";

    async function listarElementos() {
        return await paginaRelatorio
            .locator(seletorElementos)
            .evaluateAll((elementos) =>
                elementos
                    .map((elemento, index) => ({
                        index,
                        texto:
                            (
                                elemento.innerText ||
                                elemento.value ||
                                ""
                            ).trim(),
                        titulo:
                            elemento.getAttribute(
                                "title"
                            ) || "",
                        aria:
                            elemento.getAttribute(
                                "aria-label"
                            ) || "",
                        href:
                            elemento.getAttribute(
                                "href"
                            ) || "",
                        onclick:
                            elemento.getAttribute(
                                "onclick"
                            ) || ""
                    }))
                    .filter((item) => {
                        const texto =
                            `${item.texto} ${item.titulo} ${item.aria} ${item.href} ${item.onclick}`
                                .toUpperCase();

                        return (
                            texto.includes("CSV") ||
                            texto.includes("EXPORT") ||
                            texto.includes("BAIXAR") ||
                            texto.includes("DOWNLOAD") ||
                            texto.includes("EXCEL")
                        );
                    })
            );
    }

    async function tentarClique(
        elemento,
        descricao
    ) {
        try {
            if (!(await elemento.isVisible())) {
                return null;
            }

            console.log(
                "Tentando exportação:",
                descricao
            );

            const downloadPromise =
                paginaRelatorio
                    .waitForEvent("download", {
                        timeout: 4000
                    })
                    .catch(() => null);

            await elemento.click({
                timeout: 5000,
                force: true
            });

            const download =
                await downloadPromise;

            if (!download) {
                return null;
            }

            const nome =
                download.suggestedFilename();

            if (
                !String(nome)
                    .toLowerCase()
                    .endsWith(".csv")
            ) {
                return null;
            }

            return download;
        } catch (erro) {
            return null;
        }
    }

    let candidatos =
        await listarElementos();

    const elementos =
        paginaRelatorio.locator(
            seletorElementos
        );

    for (const candidato of candidatos) {
        const texto =
            `${candidato.texto} ${candidato.titulo} ${candidato.aria} ${candidato.href} ${candidato.onclick}`
                .toUpperCase();

        if (!texto.includes("CSV")) {
            continue;
        }

        const download =
            await tentarClique(
                elementos.nth(
                    candidato.index
                ),
                candidato.texto ||
                    candidato.aria ||
                    candidato.titulo ||
                    candidato.href
            );

        if (download) {
            return download;
        }
    }

    candidatos =
        await listarElementos();

    for (const candidato of candidatos) {
        const texto =
            `${candidato.texto} ${candidato.titulo} ${candidato.aria}`
                .toUpperCase();

        if (!texto.includes("EXPORT")) {
            continue;
        }

        const elementoExport =
            elementos.nth(
                candidato.index
            );

        try {
            if (
                !(await elementoExport.isVisible())
            ) {
                continue;
            }

            await elementoExport.click({
                timeout: 5000,
                force: true
            });

            await paginaRelatorio.waitForTimeout(
                500
            );
        } catch (erro) {
            continue;
        }

        const candidatosCSV =
            await paginaRelatorio
                .locator(seletorElementos)
                .evaluateAll((elementos) =>
                    elementos
                        .map((elemento, index) => ({
                            index,
                            texto:
                                (
                                    elemento.innerText ||
                                    elemento.value ||
                                    ""
                                ).trim(),
                            titulo:
                                elemento.getAttribute(
                                    "title"
                                ) || "",
                            aria:
                                elemento.getAttribute(
                                    "aria-label"
                                ) || "",
                            href:
                                elemento.getAttribute(
                                    "href"
                                ) || "",
                            onclick:
                                elemento.getAttribute(
                                    "onclick"
                                ) || ""
                        }))
                        .filter((item) => {
                            const texto =
                                `${item.texto} ${item.titulo} ${item.aria} ${item.href} ${item.onclick}`
                                    .toUpperCase();

                            return texto.includes(
                                "CSV"
                            );
                        })
                );

        for (
            const candidatoCSV of candidatosCSV
        ) {
            const download =
                await tentarClique(
                    paginaRelatorio
                        .locator(
                            seletorElementos
                        )
                        .nth(
                            candidatoCSV.index
                        ),
                    candidatoCSV.texto ||
                        candidatoCSV.aria ||
                        candidatoCSV.titulo ||
                        candidatoCSV.href
                );

            if (download) {
                return download;
            }
        }
    }

    return null;
}

// ======================================================
// 6º BIMESTRE DO FUNDEB — período anual
// ======================================================

async function selecionarSextoBimestreFundeb() {
    const seletorPeriodo =
        page.locator(
            "#ContentPlaceHolder1_ddlPeriodo"
        );

    await page.waitForFunction(
        () => {
            const select =
                document.querySelector(
                    "#ContentPlaceHolder1_ddlPeriodo"
                );

            if (!select) {
                return false;
            }

            return Array.from(
                select.options
            ).some((option) => {
                const texto =
                    (option.textContent || "")
                        .trim()
                        .toUpperCase();

                const valor =
                    String(
                        option.value || ""
                    )
                        .trim()
                        .toUpperCase();

                return (
                    texto.includes(
                        "6º BIMESTRE"
                    ) ||
                    texto.includes(
                        "6° BIMESTRE"
                    ) ||
                    texto.includes(
                        "6 BIMESTRE"
                    ) ||
                    /6\s*º?\s*BIMESTRE/.test(
                        texto
                    ) ||
                    valor === "6" ||
                    valor === "06"
                );
            });
        },
        undefined,
        { timeout: 10000 }
    );

    const opcoes =
        await seletorPeriodo
            .locator("option")
            .evaluateAll((options) =>
                options.map((option) => ({
                    value:
                        option.value,
                    texto:
                        (
                            option.textContent ||
                            ""
                        ).trim()
                }))
            );

    const sextoBimestre =
        opcoes.find((option) => {
            const texto =
                option.texto.toUpperCase();

            const valor =
                String(option.value)
                    .trim()
                    .toUpperCase();

            return (
                texto.includes(
                    "6º BIMESTRE"
                ) ||
                texto.includes(
                    "6° BIMESTRE"
                ) ||
                texto.includes(
                    "6 BIMESTRE"
                ) ||
                /6\s*º?\s*BIMESTRE/.test(
                    texto
                ) ||
                valor === "6" ||
                valor === "06"
            );
        });

    if (!sextoBimestre) {
        throw new Error(
            "6º bimestre não encontrado no seletor de período do relatório MDE."
        );
    }

    await seletorPeriodo.selectOption(
        sextoBimestre.value
    );

    await page.waitForTimeout(200);

    const valorSelecionado =
        await seletorPeriodo.inputValue({ timeout: 3000 });

    if (
        String(valorSelecionado) !==
        String(sextoBimestre.value)
    ) {
        throw new Error(
            "Não foi possível confirmar a seleção do 6º bimestre no relatório MDE."
        );
    }
}

// ======================================================
// CONSULTAR FUNDEB — coleta do FUNDEB
// ======================================================

async function consultarFundeb({
    municipioId,
    ano,
    reutilizarConfiguracao = false
}) {
    try {
        const chaveFundeb = String(municipioId);
        let configuracaoFundeb =
            fundebConfiguracoesCache.get(chaveFundeb);

        await selecionarOpcao(
            "#ContentPlaceHolder1_ddlMunicipio",
            municipioId,
            "#ContentPlaceHolder1_ddlEntidade"
        );

        if (!configuracaoFundeb) {
            const entidade = await encontrarEntidadeFundeb();

            if (!entidade) {
                return {
                    fundeb: null,
                    fundebFormatado: "-",
                    fundebOrigem: null
                };
            }

            await selecionarOpcao(
                "#ContentPlaceHolder1_ddlEntidade",
                entidade.id,
                "#ContentPlaceHolder1_ddlRelatorio"
            );

            const relatorioMDE = await encontrarRelatorioMDE();

            if (!relatorioMDE) {
                return {
                    fundeb: null,
                    fundebFormatado: "-",
                    fundebOrigem: null
                };
            }

            configuracaoFundeb = {
                entidadeId: entidade.id,
                relatorioId: relatorioMDE.id
            };

            fundebConfiguracoesCache.set(
                chaveFundeb,
                configuracaoFundeb
            );
        } else {
            await selecionarOpcao(
                "#ContentPlaceHolder1_ddlEntidade",
                configuracaoFundeb.entidadeId,
                "#ContentPlaceHolder1_ddlRelatorio"
            );
        }

        await selecionarOpcao(
            "#ContentPlaceHolder1_ddlRelatorio",
            configuracaoFundeb.relatorioId
        );

        await esperarOpcao(
            "#ContentPlaceHolder1_ddlAno",
            String(ano),
            10000
        );

        await page
            .locator("#ContentPlaceHolder1_ddlAno")
            .selectOption(String(ano));

        await page.waitForTimeout(100);

        await selecionarSextoBimestreFundeb();

        const botao = page.locator(
            "#ContentPlaceHolder1_btnConsulta"
        );

        if ((await botao.count()) === 0) {
            throw new Error(
                "Botão Consultar não encontrado para o relatório MDE."
            );
        }

        let downloadDireto = null;
        let paginaRelatorio = null;

        const resultadoEvento = await new Promise((resolve, reject) => {
            let finalizado = false;

            const limpar = () => {
                clearTimeout(timer);
                page.off("download", aoBaixar);
                context.off("page", aoAbrirPagina);
            };

            const concluir = (resultado) => {
                if (finalizado) return;
                finalizado = true;
                limpar();
                resolve(resultado);
            };

            const aoBaixar = (download) => {
                concluir({
                    download,
                    pagina: null
                });
            };

            const aoAbrirPagina = (pagina) => {
                concluir({
                    download: null,
                    pagina
                });
            };

            const timer = setTimeout(() => {
                if (finalizado) return;
                finalizado = true;
                limpar();
                reject(
                    new Error(
                        "O TCE-PR não abriu o relatório MDE nem iniciou o download dentro do tempo esperado."
                    )
                );
            }, 8000);

            page.on("download", aoBaixar);
            context.on("page", aoAbrirPagina);

            botao.click({ timeout: 5000 }).catch((erro) => {
                if (finalizado) return;
                finalizado = true;
                limpar();
                reject(erro);
            });
        });

        downloadDireto = resultadoEvento.download;
        paginaRelatorio = resultadoEvento.pagina;

        let texto = null;

        if (downloadDireto) {
            texto = await lerDownloadComoTexto(downloadDireto);
        }

        if (paginaRelatorio) {
            await paginaRelatorio
                .waitForLoadState("domcontentloaded", {
                    timeout: 10000
                })
                .catch(() => {});

            try {
                await paginaRelatorio.waitForFunction(
                    () => {
                        const texto = (
                            document.body?.innerText || ""
                        ).toUpperCase();

                        return (
                            texto.includes("FUNDEB") ||
                            texto.includes("EDUCAÇÃO") ||
                            texto.includes("EDUCACAO") ||
                            texto.includes("RECEITAS")
                        );
                    },
                    undefined,
                    { timeout: 4000 }
                );
            } catch (erro) {
                await paginaRelatorio.waitForTimeout(200);
            }

            if (!texto) {
                const downloadRelatorio =
                    await capturarDownloadDoRelatorio(
                        paginaRelatorio
                    );

                if (downloadRelatorio) {
                    texto = await lerDownloadComoTexto(
                        downloadRelatorio
                    );
                }
            }

            if (!texto) {
                const partes = [];

                try {
                    const body = await paginaRelatorio
                        .locator("body")
                        .innerText();

                    if (body) partes.push(body);
                } catch (erro) {}

                try {
                    const html = await paginaRelatorio.content();
                    if (html) partes.push(html);
                } catch (erro) {}

                for (const frame of paginaRelatorio.frames()) {
                    if (frame === paginaRelatorio.mainFrame()) {
                        continue;
                    }

                    try {
                        const textoFrame = await frame
                            .locator("body")
                            .innerText();

                        if (textoFrame) partes.push(textoFrame);
                    } catch (erro) {}

                    try {
                        const htmlFrame = await frame.content();
                        if (htmlFrame) partes.push(htmlFrame);
                    } catch (erro) {}
                }

                texto = partes.join("\n\n");
            }

            try {
                await paginaRelatorio.close();
            } catch (erro) {}
        }

        if (!texto) {
            throw new Error(
                "Não foi possível obter o conteúdo do relatório MDE."
            );
        }

        const pastaLogs = path.join(__dirname, "logs");

        if (!fs.existsSync(pastaLogs)) {
            fs.mkdirSync(pastaLogs, { recursive: true });
        }

        fs.writeFileSync(
            path.join(
                pastaLogs,
                `mde_${ano}_6bimestre.txt`
            ),
            texto,
            "utf8"
        );

        return (
            extrairFundeb(texto) || {
                fundeb: null,
                fundebFormatado: "-",
                fundebOrigem: null
            }
        );
    } catch (erro) {
        console.error(
            "ERRO AO CONSULTAR FUNDEB:",
            erro.message
        );

        return {
            fundeb: null,
            fundebFormatado: "-",
            fundebOrigem: null,
            fundebErro: erro.message
        };
    }
}

// ======================================================
// AGUARDAR RELATÓRIO LRF — espera o conteúdo
// ======================================================

async function aguardarConteudoRelatorioLRF(paginaRelatorio) {
    try {
        await paginaRelatorio.waitForFunction(
            () => {
                const texto = (document.body?.innerText || "").toUpperCase();
                return texto.includes("RECEITA CORRENTE") &&
                    texto.includes("DESPESA TOTAL COM PESSOAL");
            },
            undefined,
            { timeout: 7000 }
        );
    } catch (erro) {
        await paginaRelatorio.waitForTimeout(300);
    }
}

// ======================================================
// FUNDEB COM RETENTATIVA — nova tentativa isolada
// ======================================================

async function consultarFundebComTentativas({
    municipioId,
    ano,
    tentativas = 2
}) {
    let ultimoResultado = {
        fundeb: null,
        fundebFormatado: "-",
        fundebOrigem: null
    };

    for (let tentativa = 1; tentativa <= tentativas; tentativa++) {
        try {
            const resultado = await consultarFundeb({
                municipioId,
                ano,
                reutilizarConfiguracao: false
            });

            ultimoResultado = resultado || ultimoResultado;

            if (
                resultado &&
                resultado.fundeb !== null &&
                resultado.fundeb !== undefined &&
                !Number.isNaN(Number(resultado.fundeb))
            ) {
                return resultado;
            }
        } catch (erro) {
            ultimoResultado = {
                ...ultimoResultado,
                fundebErro: erro.message
            };
        }

        if (tentativa < tentativas) {
            await page.waitForTimeout(500);
        }
    }

    return ultimoResultado;
}

// ======================================================
// CONSULTAR UM ANO — coleta de um exercício
// ======================================================

async function consultarAno({
    municipioId,
    entidadeId,
    relatorioId,
    ano,
    periodo,
    incluirFundeb = true,
    reutilizarConfiguracao = false
}) {
    const inicioConsultaLRF = Date.now();

    if (!reutilizarConfiguracao) {
        await selecionarOpcao(
            "#ContentPlaceHolder1_ddlMunicipio",
            String(municipioId),
            "#ContentPlaceHolder1_ddlEntidade"
        );

        await selecionarOpcao(
            "#ContentPlaceHolder1_ddlEntidade",
            String(entidadeId),
            "#ContentPlaceHolder1_ddlRelatorio"
        );

        await selecionarOpcao(
            "#ContentPlaceHolder1_ddlRelatorio",
            String(relatorioId)
        );
    }

    await esperarOpcao(
        "#ContentPlaceHolder1_ddlAno",
        String(ano),
        30000
    );

    await page
        .locator(
            "#ContentPlaceHolder1_ddlAno"
        )
        .selectOption(
            String(ano)
        );

    await page.waitForTimeout(250);

    const seletorPeriodo =
        page.locator(
            "#ContentPlaceHolder1_ddlPeriodo"
        );

    if ((await seletorPeriodo.count()) > 0) {
        const existePeriodo =
            await seletorPeriodo
                .locator(
                    `option[value="${String(periodo)}"]`
                )
                .count();

        if (existePeriodo > 0) {
            await seletorPeriodo.selectOption(
                String(periodo)
            );

            await page.waitForTimeout(200);
        }
    }

    const botao =
        page.locator(
            "#ContentPlaceHolder1_btnConsulta"
        );

    if ((await botao.count()) === 0) {
        throw new Error(
            "Botão Consultar não encontrado."
        );
    }

    const novaPaginaPromise =
        context.waitForEvent("page", {
            timeout: 8000
        });

    await botao.click();

    const paginaRelatorio =
        await novaPaginaPromise;

    await paginaRelatorio
        .waitForLoadState(
            "domcontentloaded",
            {
                timeout: 10000
            }
        )
        .catch(() => {});

    await aguardarConteudoRelatorioLRF(paginaRelatorio);

    const texto =
        await paginaRelatorio
            .locator("body")
            .innerText();

    const pastaLogs =
        path.join(
            __dirname,
            "logs"
        );

    if (!fs.existsSync(pastaLogs)) {
        fs.mkdirSync(
            pastaLogs,
            { recursive: true }
        );
    }

    fs.writeFileSync(
        path.join(
            pastaLogs,
            `relatorio_${ano}_${periodo}.txt`
        ),
        texto,
        "utf8"
    );

    const dados =
        extrairDadosRelatorio(
            texto
        );

    try {
        await paginaRelatorio.close();
    } catch (erro) {}

    let fundeb = {
        fundeb: null,
        fundebFormatado: "-",
        fundebOrigem: null
    };

    console.log(`Tempo LRF ${ano}: ${Date.now() - inicioConsultaLRF} ms`);

    if (incluirFundeb) {
        const inicioConsultaFundeb = Date.now();
        fundeb =
            await consultarFundebComTentativas({
                municipioId,
                ano
            });
        console.log(`Tempo FUNDEB ${ano}: ${Date.now() - inicioConsultaFundeb} ms`);
    }

    return {
        ano,
        periodo,
        periodoFundeb:
            incluirFundeb
                ? "6º bimestre"
                : null,

        sucesso: true,

        ...dados,

        fundeb:
            fundeb?.fundeb ?? null,

        fundebFormatado:
            fundeb?.fundebFormatado || "-",

        fundebOrigem:
            fundeb?.fundebOrigem || null,

        fundebErro:
            fundeb?.fundebErro || null
    };
}

// ======================================================
// EVOLUÇÕES DA CONSULTA ÚNICA — prepara comparações
// ======================================================

async function adicionarEvolucoesConsultaUnica({
    dados,
    municipioId,
    entidadeId,
    relatorioId,
    ano,
    periodo
}) {
    const anoSelecionado = Number(ano);

    if (Number.isNaN(anoSelecionado)) {
        return dados;
    }

    const anosNecessarios = new Set([
        ANO_INICIAL,
        anoSelecionado
    ]);

    if (anoSelecionado > ANO_INICIAL) {
        anosNecessarios.add(anoSelecionado - 1);
    }

    const historico = [];

    for (const anoHistorico of Array.from(anosNecessarios).sort((a, b) => a - b)) {
        if (anoHistorico === anoSelecionado) {
            historico.push(dados);
            continue;
        }

        const item = await consultarAno({
            municipioId,
            entidadeId,
            relatorioId,
            ano: String(anoHistorico),
            periodo,
            incluirFundeb: true,
            reutilizarConfiguracao: false
        });

        historico.push(item);
    }

    for (const anoFundeb of anosNecessarios) {
        const itemFundeb = historico.find(
            item => Number(item.ano) === Number(anoFundeb)
        );

        if (!itemFundeb) {
            continue;
        }

        if (
            itemFundeb.fundeb === null ||
            itemFundeb.fundeb === undefined ||
            Number.isNaN(Number(itemFundeb.fundeb))
        ) {
            const fundeb = await consultarFundebComTentativas({
                municipioId,
                ano: String(anoFundeb)
            });

            itemFundeb.fundeb = fundeb?.fundeb ?? null;
            itemFundeb.fundebFormatado = fundeb?.fundebFormatado || "-";
            itemFundeb.fundebOrigem = fundeb?.fundebOrigem || null;
            itemFundeb.fundebErro = fundeb?.fundebErro || null;
            itemFundeb.periodoFundeb = "6º bimestre";
        }
    }

    const resultado = calcularHistorico(historico);
    const selecionado = resultado.find(
        item => Number(item.ano) === anoSelecionado
    );

    return selecionado || dados;
}

// ======================================================
// PROTEGER CONSULTAS SIMULTÂNEAS — evita conflito de sessão
// ======================================================

async function executarConsultaSegura(funcao) {
    if (consultaEmAndamento) {
        throw new Error(
            "Já existe uma consulta em andamento. Aguarde a conclusão antes de iniciar outra."
        );
    }

    consultaEmAndamento = true;

    try {
        return await funcao();
    } finally {
        consultaEmAndamento = false;
    }
}

// ======================================================
// CONSULTA ÚNICA EM PARALELO — processos isolados
// ======================================================
// CONSULTA EM PARALELO — preserva as funções de extração
// ======================================================

function iniciarWorkerConsultaUnica({
    municipioId,
    ano,
    periodo = "12"
}) {
    return new Promise((resolve, reject) => {
        const inicioWorker = Date.now();
        let finalizado = false;
        let stderr = "";

        const worker = fork(__filename, [], {
            env: {
                ...process.env,
                TCE_CONSULTA_UNICA_WORKER: "1",
                TCE_CONSULTA_UNICA_WORKER_DATA: JSON.stringify({
                    tipo: "consulta_unica",
                    municipioId,
                    ano: String(ano),
                    periodo: String(periodo)
                })
            },
            stdio: ["ignore", "pipe", "pipe", "ipc"]
        });

        worker.stdout?.on("data", () => {});
        worker.stderr?.on("data", dados => {
            stderr += dados.toString();
        });

        const timer = setTimeout(() => {
            if (finalizado) return;
            finalizado = true;

            try {
                worker.kill();
            } catch (erro) {}

            reject(
                new Error(
                    `Worker da consulta única ${ano} excedeu o tempo limite.` +
                    (stderr ? ` ${stderr.trim()}` : "")
                )
            );
        }, 120000);

        const limpar = () => clearTimeout(timer);

        worker.on("message", mensagem => {
            if (finalizado) return;

            finalizado = true;
            limpar();

            if (mensagem?.sucesso) {
                const tempo = (
                    (Date.now() - inicioWorker) / 1000
                ).toFixed(1);

                console.log(
                    `[CONSULTA ${ano}] concluída em ${tempo} s`
                );

                resolve(mensagem.dados || null);
            } else {
                reject(
                    new Error(
                        mensagem?.erro ||
                        `Worker da consulta única ${ano} falhou.`
                    )
                );
            }
        });

        worker.on("error", erro => {
            if (finalizado) return;
            finalizado = true;
            limpar();
            reject(erro);
        });

        worker.on("exit", codigo => {
            if (finalizado) return;

            finalizado = true;
            limpar();

            reject(
                new Error(
                    `Worker da consulta única ${ano} encerrou com código ${codigo}.` +
                    (stderr ? ` ${stderr.trim()}` : "")
                )
            );
        });
    });
}

async function consultarConsultaUnicaEmWorkers({
    municipioId,
    ano,
    periodo = "12"
}) {
    return await iniciarWorkerConsultaUnica({
        municipioId,
        ano,
        periodo
    });
}

// ======================================================
// SEGUNDA TENTATIVA — repete a consulta após falha
// ======================================================
// RETRY — primeira execução permanece inalterada
// ======================================================

async function consultarConsultaUnicaComTentativas({
    municipioId,
    ano,
    periodo = "12",
    tentativas = 2
}) {
    let ultimoErro = null;

    for (let tentativa = 1; tentativa <= tentativas; tentativa++) {
        try {
            return await consultarConsultaUnicaEmWorkers({
                municipioId,
                ano,
                periodo
            });
        } catch (erro) {
            ultimoErro = erro;

            if (tentativa < tentativas) {
                console.error(
                    `[CONSULTA ${ano}] tentativa ${tentativa} falhou. Recriando somente a sessão isolada da consulta...`
                );

                await new Promise(resolve => setTimeout(resolve, 500));
            }
        }
    }

    throw ultimoErro || new Error("Falha na consulta única.");
}

// ======================================================
// CONSULTA INDIVIDUAL — rota principal
// ======================================================

// A rota devolve o erro ao frontend para uma mensagem amigável ao usuário.
app.post(
    "/api/consulta",
    async (req, res) => {
        try {
            if (!pronto) {
                throw new Error(
                    "O TCE-PR ainda está sendo inicializado. Aguarde alguns segundos e tente novamente."
                );
            }

            const inicioConsultaUnica = Date.now();
            const municipioId = req.body.municipio;
            const ano = req.body.ano || String(obterAnoAtual());
            const periodo = "12";

            if (!municipioId) {
                throw new Error("Município não informado.");
            }

            const dados = await executarConsultaSegura(
                () =>
                    consultarConsultaUnicaComTentativas({
                        municipioId,
                        ano,
                        periodo,
                        tentativas: 2
                    })
            );

            console.log(`Tempo consulta única ${ano}: ${Date.now() - inicioConsultaUnica} ms`);

            res.json({
                sucesso: true,
                dados
            });
        } catch (erro) {
            console.error("ERRO NA CONSULTA:", erro.message);

            res.status(500).json({
                sucesso: false,
                erro: erro.message
            });
        }
    }
);

// ======================================================
// GET /consultar — compatibilidade
// ======================================================

app.get(
    "/consultar",
    async (req, res) => {
        try {
            const municipioId = req.query.municipio;
            const configuracao = await obterConfiguracaoConsulta(municipioId);

            const dados = await executarConsultaSegura(
                async () => {
                    const ano = req.query.ano || String(obterAnoAtual());

                    const consulta = await consultarAno({
                        municipioId,
                        entidadeId: configuracao.entidadeId,
                        relatorioId: configuracao.relatorioId,
                        ano,
                        periodo: "12",
                        incluirFundeb: true
                    });

                    return consulta;
                }
            );

            res.json({
                sucesso: true,
                dados
            });
        } catch (erro) {
            res.status(500).json({
                sucesso: false,
                erro: erro.message
            });
        }
    }
);

// ======================================================
// HISTÓRICO — dados de 2019 a 2025
// ======================================================

// O histórico permanece limitado aos sete exercícios do projeto.
app.get(
    "/historico",
    async (req, res) => {
        try {
            if (!pronto) {
                throw new Error(
                    "O TCE-PR ainda está sendo inicializado. Aguarde alguns segundos e tente novamente."
                );
            }

            const municipioId = req.query.municipio;

            if (!municipioId) {
                throw new Error("Município não informado.");
            }

            const anoInicial = ANO_INICIAL;
            const anoFinal = 2025;
            const periodo = "12";

            const resultado = await executarConsultaSegura(
                async () => {
                    const inicioHistoricoParalelo = Date.now();

                    const historico = await consultarHistoricoEmWorkers({
                        municipioId,
                        anos: Array.from(
                            { length: anoFinal - anoInicial + 1 },
                            (_, indice) => String(anoInicial + indice)
                        ),
                        quantidadeWorkers: 8
                    });

                    console.log(
                        `Tempo HISTÓRICO PARALELO: ${Date.now() - inicioHistoricoParalelo} ms`
                    );

                    return calcularHistorico(historico).filter(
                        item => Number(item.ano) >= anoInicial
                    );
                }
            );

            historicoExportCache.set(String(municipioId), {
                dados: resultado,
                municipioNome:
                    municipiosCache.find(
                        municipio =>
                            String(municipio.id) === String(municipioId)
                    )?.nome || "Municipio"
            });

            res.json({
                sucesso: true,
                anoInicial,
                anoFinal,
                periodo,
                dados: resultado
            });
        } catch (erro) {
            console.error("ERRO NA CONSULTA HISTÓRICA:", erro.message);

            res.status(500).json({
                sucesso: false,
                erro: erro.message
            });
        }
    }
);


// ======================================================
// EXPORTAÇÃO EXCEL — geração do arquivo
// ======================================================

function corPercentualExcel(percentual) {
    const valor = Number(percentual);

    if (Number.isNaN(valor)) return null;

    if (valor >= 54) {
        return ["FFD71920", "FFFFFFFF"];
    }

    if (valor >= 51.3) {
        return ["FFF59E0B", "FFFFFFFF"];
    }

    if (valor >= 48.6) {
        return ["FFEAD800", "FF2C2C00"];
    }

    return ["FF55C83B", "FFFFFFFF"];
}

function formatarEvolucaoExcel(valor) {
    if (
        valor === null ||
        valor === undefined ||
        Number.isNaN(Number(valor))
    ) {
        return "-";
    }

    const numero = Number(valor);
    const sinal = numero > 0 ? "+" : "";

    return (
        sinal +
        numero.toFixed(2).replace(".", ",") +
        "%"
    );
}

function calcularEvolucaoExcel(atual, anterior) {
    if (
        atual === null ||
        atual === undefined ||
        anterior === null ||
        anterior === undefined ||
        Number.isNaN(Number(atual)) ||
        Number.isNaN(Number(anterior)) ||
        Number(anterior) === 0
    ) {
        return null;
    }

    return (
        ((Number(atual) - Number(anterior)) /
            Number(anterior)) *
        100
    );
}

function calcularAcumuladaExcel(atual, base) {
    if (
        atual === null ||
        atual === undefined ||
        base === null ||
        base === undefined ||
        Number.isNaN(Number(atual)) ||
        Number.isNaN(Number(base)) ||
        Number(base) === 0
    ) {
        return null;
    }

    return (
        ((Number(atual) - Number(base)) /
            Number(base)) *
        100
    );
}

function montarDadosExcel(historico) {
    const lista = Array.isArray(historico)
        ? historico
        : [];

    const anos = lista.map(
        item => String(item.ano)
    );

    const base =
        lista.find(
            item => Number(item.ano) === ANO_INICIAL
        ) || lista[0];

    const linhas = [
        ["Indicador", "Tipo", ...anos]
    ];

    const adicionarLinha = (
        indicador,
        tipo,
        valores
    ) => {
        linhas.push([
            indicador,
            tipo,
            ...valores
        ]);
    };

    const evolucaoAnual = (
        campo
    ) =>
        lista.map(
            (item, indice) =>
                formatarEvolucaoExcel(
                    indice === 0
                        ? null
                        : calcularEvolucaoExcel(
                              item[campo],
                              lista[indice - 1]?.[campo]
                          )
                )
        );

    const evolucaoAcumulada = (
        campo
    ) =>
        lista.map(
            (item, indice) =>
                formatarEvolucaoExcel(
                    indice === 0
                        ? null
                        : calcularAcumuladaExcel(
                              item[campo],
                              base?.[campo]
                          )
                )
        );

    adicionarLinha(
        "RREO — Receita Corrente Líquida Ajustada",
        "Valor",
        lista.map(
            item =>
                item.rclAjustadaFormatada ||
                formatarBR(item.rclAjustada)
        )
    );

    adicionarLinha(
        "Evolução anual",
        "%",
        evolucaoAnual("rclAjustada")
    );

    adicionarLinha(
        "Evolução acumulada",
        "%",
        evolucaoAcumulada("rclAjustada")
    );

    adicionarLinha(
        "RGF — Despesa Total com Pessoal",
        "Valor",
        lista.map(
            item =>
                item.dtpFormatado ||
                formatarBR(item.dtp)
        )
    );

    adicionarLinha(
        "Evolução anual",
        "%",
        evolucaoAnual("dtp")
    );

    adicionarLinha(
        "Evolução acumulada",
        "%",
        evolucaoAcumulada("dtp")
    );

    adicionarLinha(
        "RGF — Despesa Total com Pessoal %",
        "Valor",
        lista.map(
            item =>
                item.percentualFormatado ||
                `${Number(item.percentual || 0).toFixed(2)}%`
        )
    );

    adicionarLinha(
        "RREO — Transferências FUNDEB",
        "Valor",
        lista.map(
            item =>
                item.fundebFormatado ||
                formatarBR(item.fundeb)
        )
    );

    adicionarLinha(
        "Evolução anual",
        "%",
        evolucaoAnual("fundeb")
    );

    adicionarLinha(
        "Evolução acumulada",
        "%",
        evolucaoAcumulada("fundeb")
    );

    return {
        lista,
        anos,
        linhas
    };
}

async function criarImagemGraficoExcel(lista) {
    const largura = 1000;
    const altura = 520;
    const esquerda = 105;
    const direita = 35;
    const topo = 92;
    const base = 62;
    const larguraGrafico =
        largura - esquerda - direita;
    const alturaGrafico =
        altura - topo - base;

    const maximo = Math.max(
        1,
        ...lista.flatMap(item => [
            Number(item.rclAjustada) || 0,
            Number(item.fundeb) || 0,
            Number(item.dtp) || 0
        ])
    );

    const teto = maximo * 1.12;
    const quantidade =
        Math.max(lista.length, 1);
    const grupo =
        larguraGrafico / quantidade;
    const barra = Math.min(
        30,
        grupo * 0.18
    );
    const espacamento = barra * 0.22;

    const y = valor =>
        topo +
        alturaGrafico -
        ((Number(valor) || 0) / teto) *
            alturaGrafico;

    const grade = [];
    const eixo = [];

    for (let i = 0; i <= 5; i++) {
        const valor =
            (teto / 5) * i;
        const yy = y(valor);

        grade.push(`
            <line
                x1="${esquerda}"
                y1="${yy}"
                x2="${largura - direita}"
                y2="${yy}"
                stroke="#dfe7ef"
                stroke-width="1"
            />
        `);

        eixo.push(`
            <text
                x="${esquerda - 14}"
                y="${yy + 5}"
                text-anchor="end"
                font-family="Arial"
                font-size="18"
                fill="#617286"
            >${formatarBR(Math.round(valor))}</text>
        `);
    }

    const barras = [];
    const anos = [];

    lista.forEach(
        (item, indice) => {
            const centro =
                esquerda +
                grupo * indice +
                grupo / 2;

            [
                [
                    item.rclAjustada,
                    "#2477a8",
                    -barra - espacamento
                ],
                [
                    item.fundeb,
                    "#30964f",
                    0
                ],
                [
                    item.dtp,
                    "#d44949",
                    barra + espacamento
                ]
            ].forEach(
                ([valor, cor, deslocamento]) => {
                    const topoBarra = y(valor);
                    const alturaBarra =
                        topo +
                        alturaGrafico -
                        topoBarra;

                    barras.push(`
                        <rect
                            x="${centro + deslocamento - barra / 2}"
                            y="${topoBarra}"
                            width="${barra}"
                            height="${Math.max(0, alturaBarra)}"
                            fill="${cor}"
                        />
                    `);
                }
            );

            anos.push(`
                <text
                    x="${centro}"
                    y="${altura - 48}"
                    text-anchor="middle"
                    font-family="Arial"
                    font-size="19"
                    font-weight="700"
                    fill="#617286"
                >${String(item.ano)}</text>
            `);
        }
    );

    const svg = `
        <svg
            xmlns="http://www.w3.org/2000/svg"
            width="${largura}"
            height="${altura}"
        >
            <rect
                width="100%"
                height="100%"
                fill="#ffffff"
            />

            <text
                x="${esquerda}"
                y="34"
                font-family="Arial"
                font-size="22"
                font-weight="700"
                fill="#122b42"
            >Evolução dos principais indicadores financeiros</text>

            <rect x="105" y="46" width="24" height="12" fill="#2477a8"/>
            <text x="139" y="58" font-family="Arial" font-size="16" fill="#33465d">RCL Ajustada</text>

            <rect x="300" y="46" width="24" height="12" fill="#30964f"/>
            <text x="334" y="58" font-family="Arial" font-size="16" fill="#33465d">FUNDEB</text>

            <rect x="470" y="46" width="24" height="12" fill="#d44949"/>
            <text x="504" y="58" font-family="Arial" font-size="16" fill="#33465d">DTP</text>

            ${grade.join("")}
            ${eixo.join("")}

            <line
                x1="${esquerda}"
                y1="${topo + alturaGrafico}"
                x2="${largura - direita}"
                y2="${topo + alturaGrafico}"
                stroke="#b9c6d3"
                stroke-width="1"
            />

            ${barras.join("")}
            ${anos.join("")}
        </svg>
    `;

    return sharp(
        Buffer.from(svg)
    )
        .png()
        .toBuffer();
}

async function gerarExcelFormatado(
    historico,
    municipioNome
) {
    const {
        lista,
        anos,
        linhas
    } = montarDadosExcel(historico);

    const workbook =
        new ExcelJS.Workbook();

    workbook.creator =
        "Automação TCE-PR";
    workbook.title =
        "Análise TCE-PR";
    workbook.subject =
        "Histórico financeiro municipal";
    workbook.created =
        new Date();

    const worksheet =
        workbook.addWorksheet(
            "Histórico 2019-2025",
            {
                views: [
                    {
                        state: "frozen",
                        ySplit: 5,
                        xSplit: 2,
                        zoomScale: 90
                    }
                ]
            }
        );

    const ultimaColuna =
        String.fromCharCode(
            64 + anos.length + 2
        );

    worksheet.mergeCells(
        `A1:${ultimaColuna}1`
    );
    worksheet.mergeCells(
        `A2:${ultimaColuna}2`
    );
    worksheet.mergeCells(
        `A3:${ultimaColuna}3`
    );

    worksheet.getCell("A1").value =
        "ANÁLISE TCE-PR — HISTÓRICO 2019–2025";

    worksheet.getCell("A1").font = {
        name: "Arial",
        size: 18,
        bold: true,
        color: {
            argb: "FFFFFFFF"
        }
    };

    worksheet.getCell("A1").fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: {
            argb: "FF123B5D"
        }
    };

    worksheet.getCell("A1").alignment = {
        vertical: "middle",
        horizontal: "left"
    };

    worksheet.getRow(1).height = 34;

    worksheet.getCell("A2").value =
        `Município: ${municipioNome} • Período: dezembro • Anos: ${anos.join(", ")}`;

    worksheet.getCell("A2").font = {
        name: "Arial",
        size: 11,
        bold: true,
        color: {
            argb: "FF176DA8"
        }
    };

    worksheet.getCell("A2").fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: {
            argb: "FFEAF5FF"
        }
    };

    worksheet.getCell("A2").alignment = {
        vertical: "middle",
        horizontal: "left"
    };

    worksheet.getRow(2).height = 24;

    worksheet.getCell("A3").value =
        "Histórico financeiro municipal com evolução anual, evolução acumulada, percentual da DTP e referência visual dos limites da LRF.";

    worksheet.getCell("A3").font = {
        name: "Arial",
        size: 10,
        italic: true,
        color: {
            argb: "FF66788C"
        }
    };

    worksheet.getCell("A3").alignment = {
        vertical: "middle",
        horizontal: "left",
        wrapText: true
    };

    worksheet.getRow(3).height = 30;

    linhas.forEach(
        (linha, indice) => {
            const row =
                worksheet.addRow(linha);

            row.height =
                indice === 0
                    ? 32
                    : 27;

            row.eachCell(
                (cell, coluna) => {
                    cell.border = {
                        top: {
                            style: "thin",
                            color: {
                                argb: "FFDCE5ED"
                            }
                        },
                        left: {
                            style: "thin",
                            color: {
                                argb: "FFDCE5ED"
                            }
                        },
                        bottom: {
                            style: "thin",
                            color: {
                                argb: "FFDCE5ED"
                            }
                        },
                        right: {
                            style: "thin",
                            color: {
                                argb: "FFDCE5ED"
                            }
                        }
                    };

                    cell.alignment = {
                        vertical: "middle",
                        horizontal:
                            coluna <= 2
                                ? "left"
                                : "right",
                        wrapText: true
                    };

                    if (indice === 0) {
                        cell.fill = {
                            type: "pattern",
                            pattern: "solid",
                            fgColor: {
                                argb: "FF176DA8"
                            }
                        };

                        cell.font = {
                            name: "Arial",
                            size: 11,
                            bold: true,
                            color: {
                                argb: "FFFFFFFF"
                            }
                        };

                        cell.alignment = {
                            vertical: "middle",
                            horizontal:
                                coluna <= 2
                                    ? "left"
                                    : "center"
                        };
                    } else {
                        cell.font = {
                            name: "Arial",
                            size: 10,
                            bold: coluna <= 2,
                            color: {
                                argb: "FF162B40"
                            }
                        };

                        if (coluna <= 2) {
                            cell.fill = {
                                type: "pattern",
                                pattern: "solid",
                                fgColor: {
                                    argb:
                                        indice % 2 === 0
                                            ? "FFF8FAFC"
                                            : "FFFFFFFF"
                                }
                            };
                        }
                    }
                }
            );

            if (
                indice > 0 &&
                linha[1] === "%"
            ) {
                for (
                    let coluna = 3;
                    coluna <= linha.length;
                    coluna++
                ) {
                    const texto =
                        row.getCell(coluna).value;

                    if (
                        typeof texto !== "string" ||
                        texto === "-"
                    ) {
                        continue;
                    }

                    const valor =
                        Number(
                            texto
                                .replace("+", "")
                                .replace("%", "")
                                .replace(",", ".")
                        );

                    if (Number.isNaN(valor)) {
                        continue;
                    }

                    const positivo =
                        valor > 0;
                    const negativo =
                        valor < 0;

                    row.getCell(coluna).fill = {
                        type: "pattern",
                        pattern: "solid",
                        fgColor: {
                            argb: positivo
                                ? "FFE8F5E9"
                                : negativo
                                    ? "FFFFEEEE"
                                    : "FFF3F6FA"
                        }
                    };

                    row.getCell(coluna).font = {
                        name: "Arial",
                        size: 10,
                        bold: true,
                        color: {
                            argb: positivo
                                ? "FF217A19"
                                : negativo
                                    ? "FFD71920"
                                    : "FF66788C"
                        }
                    };
                }
            }

            if (
                indice > 0 &&
                linha[0] ===
                    "RGF — Despesa Total com Pessoal %"
            ) {
                for (
                    let coluna = 3;
                    coluna <= linha.length;
                    coluna++
                ) {
                    const texto =
                        row.getCell(coluna).value;

                    const valor =
                        Number(
                            String(texto)
                                .replace("%", "")
                                .replace(",", ".")
                        );

                    const cor =
                        corPercentualExcel(valor);

                    if (!cor) continue;

                    row.getCell(coluna).fill = {
                        type: "pattern",
                        pattern: "solid",
                        fgColor: {
                            argb: cor[0]
                        }
                    };

                    row.getCell(coluna).font = {
                        name: "Arial",
                        size: 10,
                        bold: true,
                        color: {
                            argb: cor[1]
                        }
                    };

                    row.getCell(coluna).alignment = {
                        vertical: "middle",
                        horizontal: "center"
                    };
                }
            }
        }
    );

    worksheet.getColumn(1).width = 52;
    worksheet.getColumn(2).width = 14;

    for (
        let coluna = 3;
        coluna <= Math.max(anos.length + 2, 12);
        coluna++
    ) {
        worksheet.getColumn(coluna).width = 19;
    }

    worksheet.autoFilter = {
        from: {
            row: 5,
            column: 1
        },
        to: {
            row: 5,
            column: anos.length + 2
        }
    };

    worksheet.pageSetup = {
        orientation: "landscape",
        fitToPage: true,
        fitToWidth: 1,
        fitToHeight: 0,
        paperSize:
            worksheet.PAPERSIZE_A4
    };

    worksheet.getColumn(10).width = 16;
    worksheet.getColumn(11).width = 27;
    worksheet.getColumn(12).width = 3;

    worksheet.mergeCells("J5:K5");

    const tituloLegenda =
        worksheet.getCell("J5");

    tituloLegenda.value =
        "Legenda — Percentual DTP";

    tituloLegenda.font = {
        name: "Arial",
        size: 11,
        bold: true,
        color: {
            argb: "FFFFFFFF"
        }
    };

    tituloLegenda.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: {
            argb: "FF176DA8"
        }
    };

    tituloLegenda.alignment = {
        vertical: "middle",
        horizontal: "left"
    };

    worksheet.getRow(5).height = 30;

    [
        [
            "Até 48,6%",
            "Situação regular",
            "FF55C83B"
        ],
        [
            "48,6% a 51,3%",
            "Alerta",
            "FFEAD800"
        ],
        [
            "51,3% a 54%",
            "Limite prudencial",
            "FFF59E0B"
        ],
        [
            "Acima de 54%",
            "Acima do limite legal",
            "FFD71920"
        ]
    ].forEach(
        (item, indice) => {
            const rowNumero = 6 + indice;
            const row = worksheet.getRow(rowNumero);

            row.height = 24;

            const faixa = row.getCell(10);
            const descricao = row.getCell(11);

            faixa.value = item[0];
            descricao.value = item[1];

            faixa.fill = {
                type: "pattern",
                pattern: "solid",
                fgColor: {
                    argb: item[2]
                }
            };

            faixa.font = {
                name: "Arial",
                size: 9,
                bold: true,
                color: {
                    argb:
                        item[2] === "FFEAD800"
                            ? "FF2C2C00"
                            : "FFFFFFFF"
                }
            };

            descricao.font = {
                name: "Arial",
                size: 9,
                bold: true,
                color: {
                    argb: "FF33465D"
                }
            };

            faixa.alignment = {
                vertical: "middle",
                horizontal: "center",
                wrapText: true
            };

            descricao.alignment = {
                vertical: "middle",
                horizontal: "left",
                wrapText: true
            };

            [faixa, descricao].forEach(cell => {
                cell.border = {
                    top: {
                        style: "thin",
                        color: {
                            argb: "FFDCE5ED"
                        }
                    },
                    left: {
                        style: "thin",
                        color: {
                            argb: "FFDCE5ED"
                        }
                    },
                    bottom: {
                        style: "thin",
                        color: {
                            argb: "FFDCE5ED"
                        }
                    },
                    right: {
                        style: "thin",
                        color: {
                            argb: "FFDCE5ED"
                        }
                    }
                };
            });
        }
    );

    const grafico =
        workbook.addWorksheet(
            "Gráfico"
        );

    grafico.views = [
        {
            showGridLines: false,
            zoomScale: 85
        }
    ];

    grafico.getCell("A1").value =
        `Evolução dos principais indicadores — ${municipioNome}`;

    grafico.getCell("A1").font = {
        name: "Arial",
        size: 18,
        bold: true,
        color: {
            argb: "FF122B42"
        }
    };

    const imagem =
        await criarImagemGraficoExcel(
            lista
        );

    grafico.getRow(1).height = 28;

    for (let coluna = 1; coluna <= 12; coluna++) {
        grafico.getColumn(coluna).width = 10;
    }

    const imagemId =
        workbook.addImage({
            buffer: imagem,
            extension: "png"
        });

    grafico.addImage(
        imagemId,
        {
            tl: {
                col: 0,
                row: 2
            },
            ext: {
                width: 900,
                height: 468
            }
        }
    );

    return workbook.xlsx.writeBuffer();
}

app.get(
    "/api/exportar-excel",
    async (req, res) => {
        try {
            const municipioId =
                req.query.municipio;

            if (!municipioId) {
                return res.status(400).json({
                    sucesso: false,
                    erro: "Município não informado."
                });
            }

            const cache =
                historicoExportCache.get(
                    String(municipioId)
                );

            if (
                !cache ||
                !Array.isArray(cache.dados) ||
                cache.dados.length === 0
            ) {
                return res.status(400).json({
                    sucesso: false,
                    erro:
                        "Consulte o histórico antes de exportar para o Excel."
                });
            }

            const arquivo =
                await gerarExcelFormatado(
                    cache.dados,
                    cache.municipioNome
                );

            const nomeArquivo =
                `Historico_${cache.municipioNome
                    .normalize("NFD")
                    .replace(/[\u0300-\u036f]/g, "")
                    .replace(/[^a-zA-Z0-9 ]/g, "")
                    .replace(/\s+/g, "_")}_2019-2025.xlsx`;

            res.setHeader(
                "Content-Type",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            );

            res.setHeader(
                "Content-Disposition",
                `attachment; filename="${nomeArquivo}"`
            );

            res.end(arquivo);
        } catch (erro) {
            console.error(
                "ERRO AO GERAR EXCEL:",
                erro.message
            );

            res.status(500).json({
                sucesso: false,
                erro:
                    "Não foi possível gerar o arquivo Excel formatado."
            });
        }
    }
);

// ======================================================
// HISTÓRICO PARALELO — workers por ano
// ======================================================

function iniciarWorkerHistorico({ municipioId, anosDoWorker, indice, totalWorkers }) {
    return new Promise((resolve, reject) => {
        const worker = fork(__filename, [], {
            env: {
                ...process.env,
                TCE_HISTORICO_WORKER: "1",
                TCE_HISTORICO_WORKER_DATA: JSON.stringify({
                    municipioId,
                    anos: anosDoWorker
                })
            },
            stdio: ["ignore", "pipe", "pipe", "ipc"]
        });

        workersHistoricoAtivos++;

        let finalizado = false;
        let stdout = "";
        let stderr = "";

        const saidaWorker = new Promise(resolveSaida => {
            worker.once("exit", codigo => {
                workersHistoricoAtivos = Math.max(0, workersHistoricoAtivos - 1);

                if (workersHistoricoAtivos === 0) {
                    const aguardando = aguardandoWorkersHistorico;
                    aguardandoWorkersHistorico = [];
                    aguardando.forEach(resolveAguardando => resolveAguardando());
                }

                resolveSaida(codigo);
            });
        });

        worker.stdout.on("data", dados => {
            stdout += dados.toString();
        });

        worker.stderr.on("data", dados => {
            stderr += dados.toString();
        });

        worker.on("message", mensagem => {
            if (finalizado) return;

            if (mensagem?.sucesso) {
                finalizado = true;

                console.log(
                    `Worker histórico ${indice + 1}/${totalWorkers} terminou: ${anosDoWorker.join(", ")}`
                );

                resolve(mensagem.dados || []);
                return;
            }

            finalizado = true;

            const erro = new Error(
                mensagem?.erro ||
                `Worker histórico ${indice + 1} falhou.`
            );

            saidaWorker.then(() => reject(erro));
        });

        worker.on("error", erro => {
            if (finalizado) return;
            finalizado = true;
            saidaWorker.then(() => reject(erro));
        });

        worker.on("exit", codigo => {
            if (finalizado) return;

            finalizado = true;

            reject(
                new Error(
                    `Worker histórico ${indice + 1} encerrou com código ${codigo}.` +
                    (stderr ? ` ${stderr.trim()}` : "") +
                    (stdout ? ` ${stdout.trim()}` : "")
                )
            );
        });
    });
}

function aguardarFinalizacaoWorkersHistorico() {
    if (workersHistoricoAtivos === 0) {
        return Promise.resolve();
    }

    return new Promise(resolve => {
        aguardandoWorkersHistorico.push(resolve);
    });
}

async function executarWorkerHistoricoComTentativas({
    municipioId,
    anosDoWorker,
    indice,
    totalWorkers,
    tentativas = 2
}) {
    let ultimoErro = null;

    for (let tentativa = 1; tentativa <= tentativas; tentativa++) {
        try {
            return await iniciarWorkerHistorico({
                municipioId,
                anosDoWorker,
                indice,
                totalWorkers
            });
        } catch (erro) {
            ultimoErro = erro;

            if (tentativa < tentativas) {
                console.error(
                    `Worker histórico ${indice + 1}/${totalWorkers} (${anosDoWorker.join(", ")}) falhou na tentativa ${tentativa}. Repetindo somente este ano...`
                );

                await new Promise(resolve => setTimeout(resolve, 1000));
            }
        }
    }

    throw ultimoErro ||
        new Error(
            `Worker histórico ${indice + 1}/${totalWorkers} falhou após ${tentativas} tentativas.`
        );
}

async function consultarHistoricoEmWorkers({ municipioId, anos, quantidadeWorkers = 8 }) {
    const blocos = anos
        .map(ano => [ano])
        .slice(0, quantidadeWorkers);

    const resultados = await Promise.allSettled(
        blocos.map((anosDoWorker, indice) =>
            executarWorkerHistoricoComTentativas({
                municipioId,
                anosDoWorker,
                indice,
                totalWorkers: blocos.length,
                tentativas: 2
            })
        )
    );

    limpezaHistoricoPendente = aguardarFinalizacaoWorkersHistorico();

    const falhas = resultados.filter(
        resultado => resultado.status === "rejected"
    );

    if (falhas.length > 0) {
        throw new Error(
            falhas
                .map(resultado => resultado.reason?.message || "Falha no worker histórico.")
                .join(" | ")
        );
    }

    return resultados
        .flatMap(resultado => resultado.value || [])
        .sort((a, b) => Number(a.ano) - Number(b.ano));
}

// ======================================================
// FECHAR PLAYWRIGHT — encerramento do navegador
// ======================================================

async function fecharTCE() {
    try {
        if (context) {
            await context.close();
        }
    } catch (erro) {}

    try {
        if (browser) {
            await browser.close();
        }
    } catch (erro) {}

    context = null;
    browser = null;
    page = null;
    pronto = false;
}

// ======================================================
// SERVIDOR — inicialização e rotas
// ======================================================

if (process.env.TCE_CONSULTA_UNICA_WORKER === "1") {
    (async () => {
        try {
            const dadosWorker = JSON.parse(
                process.env.TCE_CONSULTA_UNICA_WORKER_DATA || "{}"
            );

            await iniciarTCE();

            if (dadosWorker.tipo === "consulta_unica") {
                const inicioConsulta = Date.now();

                const configuracao = await obterConfiguracaoConsulta(
                    dadosWorker.municipioId
                );

                const lrf = await consultarAno({
                    municipioId: dadosWorker.municipioId,
                    entidadeId: configuracao.entidadeId,
                    relatorioId: configuracao.relatorioId,
                    ano: String(dadosWorker.ano),
                    periodo: String(dadosWorker.periodo || "12"),
                    incluirFundeb: false,
                    reutilizarConfiguracao: false
                });

                if (!lrf) {
                    throw new Error("A consulta LRF não retornou dados.");
                }

                const fundeb = await consultarFundebComTentativas({
                    municipioId: dadosWorker.municipioId,
                    ano: String(dadosWorker.ano)
                });

                if (
                    !fundeb ||
                    fundeb.fundeb === null ||
                    fundeb.fundeb === undefined ||
                    Number.isNaN(Number(fundeb.fundeb))
                ) {
                    throw new Error(
                        `FUNDEB ${dadosWorker.ano} não foi obtido após as tentativas internas.`
                    );
                }

                const dados = {
                    ...lrf,
                    fundeb: fundeb.fundeb,
                    fundebFormatado: fundeb.fundebFormatado || "-",
                    fundebOrigem: fundeb.fundebOrigem || null,
                    fundebErro: fundeb.fundebErro || null,
                    periodoFundeb: "6º bimestre"
                };

                console.log(
                    `Consulta única isolada ${dadosWorker.ano} concluída em ${((Date.now() - inicioConsulta) / 1000).toFixed(1)} s`
                );

                if (typeof process.send === "function") {
                    process.send({
                        sucesso: true,
                        dados
                    });
                }
            } else if (dadosWorker.tipo === "lrf") {
                const configuracao = await obterConfiguracaoConsulta(
                    dadosWorker.municipioId
                );

                const dados = await consultarAno({
                    municipioId: dadosWorker.municipioId,
                    entidadeId: configuracao.entidadeId,
                    relatorioId: configuracao.relatorioId,
                    ano: String(dadosWorker.ano),
                    periodo: String(dadosWorker.periodo || "12"),
                    incluirFundeb: false,
                    reutilizarConfiguracao: true
                });

                if (typeof process.send === "function") {
                    process.send({
                        sucesso: true,
                        dados
                    });
                }
            } else if (dadosWorker.tipo === "fundeb") {
                const dados = await consultarFundebComTentativas({
                    municipioId: dadosWorker.municipioId,
                    ano: String(dadosWorker.ano)
                });

                if (typeof process.send === "function") {
                    process.send({
                        sucesso: true,
                        dados
                    });
                }
            } else {
                throw new Error("Tipo de worker de consulta única inválido.");
            }
        } catch (erro) {
            if (typeof process.send === "function") {
                process.send({
                    sucesso: false,
                    erro: erro.message
                });
            }
        } finally {
            await fecharTCE();
            process.exit(0);
        }
    })();
} else if (process.env.TCE_HISTORICO_WORKER === "1") {
    (async () => {
        try {
            const dadosWorker = JSON.parse(
                process.env.TCE_HISTORICO_WORKER_DATA || "{}"
            );

            await iniciarTCE();

            let configuracao = null;
            let erroConfiguracao = null;

            for (let tentativaConfiguracao = 1; tentativaConfiguracao <= 2; tentativaConfiguracao++) {
                try {
                    configuracao = await obterConfiguracaoConsulta(
                        dadosWorker.municipioId
                    );
                    erroConfiguracao = null;
                    break;
                } catch (erro) {
                    erroConfiguracao = erro;

                    if (tentativaConfiguracao < 2) {
                        await selecionarOpcao(
                            "#ContentPlaceHolder1_ddlMunicipio",
                            String(dadosWorker.municipioId),
                            "#ContentPlaceHolder1_ddlEntidade"
                        );
                    }
                }
            }

            if (!configuracao) {
                throw erroConfiguracao ||
                    new Error("Não foi possível carregar a configuração do município.");
            }

            const resultados = [];

            for (const ano of dadosWorker.anos || []) {
                const inicioAno = Date.now();

                const dados = await consultarAno({
                    municipioId: dadosWorker.municipioId,
                    entidadeId: configuracao.entidadeId,
                    relatorioId: configuracao.relatorioId,
                    ano: String(ano),
                    periodo: "12",
                    incluirFundeb: true,
                    reutilizarConfiguracao: false
                });

                if (
                    !dados ||
                    dados.fundeb === null ||
                    dados.fundeb === undefined ||
                    Number.isNaN(Number(dados.fundeb))
                ) {
                    throw new Error(
                        `FUNDEB ${ano} não foi obtido. ${dados?.fundebErro || "Nova tentativa necessária."}`
                    );
                }

                resultados.push(dados);

                console.log(
                    `Worker: ${ano} concluído em ${((Date.now() - inicioAno) / 1000).toFixed(1)} s`
                );
            }

            if (typeof process.send === "function") {
                process.send({
                    sucesso: true,
                    dados: resultados
                });
            }
        } catch (erro) {
            if (typeof process.send === "function") {
                process.send({
                    sucesso: false,
                    erro: erro.message
                });
            }
        } finally {
            await fecharTCE();
            process.exit(0);
        }
    })();
} else {
    app.listen(
        PORT,
        async () => {
            console.log("");
            console.log("=================================");
            console.log(
                `Servidor: http://localhost:${PORT}`
            );
            console.log(
                "Ano atual:",
                obterAnoAtual()
            );
            console.log(
                "Mês atual:",
                obterMesAtual()
            );
            console.log(
                "Anos disponíveis:",
                gerarAnos()
            );
            console.log("=================================");

            try {
                await iniciarTCE();
            } catch (erro) {
                console.error(
                    "ERRO AO INICIAR TCE:",
                    erro.message
                );
            }
        }
    );
}

// ======================================================
// ENCERRAMENTO LIMPO — liberação dos recursos
// ======================================================

process.on(
    "SIGINT",
    async () => {
        console.log(
            "\nEncerrando servidor..."
        );

        await fecharTCE();

        process.exit(0);
    }
);

process.on(
    "SIGTERM",
    async () => {
        await fecharTCE();
        process.exit(0);
    }
);
