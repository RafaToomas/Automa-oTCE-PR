const municipioSelect =
        document.getElementById("municipio");

    const anoSelect =
        document.getElementById("ano");

    const btnConsultar =
        document.getElementById("btnConsultar");

    const btnHistorico =
        document.getElementById("btnHistorico");

    const status =
        document.getElementById("status");

    const mensagem =
        document.getElementById("mensagem");

    const erroBox =
        document.getElementById("erroBox");

    const resultadoUnico =
        document.getElementById("resultadoUnico");

    const resultadoHistorico =
        document.getElementById("resultadoHistorico");

    const historicoTabela =
        document.getElementById("tabelaHistorico");

    const historicoTabelaContainer =
        document.getElementById("historicoTabelaContainer");

    const historicoGrafico =
        document.getElementById("historicoGrafico");

    const btnAlternarHistorico =
        document.getElementById("btnAlternarHistorico");

    const btnExportarExcel =
        document.getElementById("btnExportarExcel");

    const loadingOverlay =
        document.getElementById("loadingOverlay");

    const loadingMensagem =
        document.getElementById("loadingMensagem");


    let dadosHistoricoAtual = [];

    let graficoHistorico = null;

    let intervaloLoading = null;


    function escaparHTML(valor) {

        return String(valor ?? "")
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#039;");
    }


    function mostrarErro(texto) {

        erroBox.textContent =
            texto || "Ocorreu um erro.";

        erroBox.classList.add("visivel");
    }


    function esconderErro() {

        erroBox.textContent = "";

        erroBox.classList.remove("visivel");
    }


    function definirCarregando(texto) {

        status.className =
            "status carregando";

        status.textContent =
            texto;
    }


    function definirSucesso(texto) {

        status.className =
            "status";

        status.textContent =
            texto;
    }


    function definirErroStatus(texto) {

        status.className =
            "status erro";

        status.textContent =
            texto;
    }


    function iniciarLoading(tipo) {

        loadingOverlay.classList.add("visivel");

        const mensagensUnica = [

            "Acessando o TCE-PR...",

            "Localizando o município...",

            "Consultando dados da RCL e DTP...",

            "Obtendo informações do FUNDEB...",

            "Organizando os resultados..."

        ];


        const mensagensHistorico = [

            "Acessando o TCE-PR...",

            "Localizando o município...",

            "Consultando os dados históricos...",

            "Obtendo RCL e DTP dos anos selecionados...",

            "Buscando informações do FUNDEB...",

            "Organizando o histórico..."

        ];


        const mensagens =
            tipo === "historico"
                ? mensagensHistorico
                : mensagensUnica;


        let indice = 0;

        loadingMensagem.textContent =
            mensagens[0];


        intervaloLoading =
            setInterval(() => {

                indice =
                    (indice + 1) %
                    mensagens.length;

                loadingMensagem.textContent =
                    mensagens[indice];

            }, 2500);
    }


    function finalizarLoading() {

        if (intervaloLoading) {

            clearInterval(intervaloLoading);

            intervaloLoading = null;
        }

        loadingOverlay.classList.remove("visivel");
    }


    function bloquearBotoes() {

        btnConsultar.disabled = true;

        btnHistorico.disabled = true;
    }


    function desbloquearBotoes() {

        btnConsultar.disabled = false;

        btnHistorico.disabled = false;
    }


    function formatarMoeda(valor) {

        if (
            valor === null ||
            valor === undefined ||
            valor === "" ||
            Number.isNaN(Number(valor))
        ) {
            return "-";
        }

        return Number(valor).toLocaleString(
            "pt-BR",
            {
                minimumFractionDigits: 2,
                maximumFractionDigits: 2
            }
        );
    }


    function formatarPercentual(valor) {

        if (
            valor === null ||
            valor === undefined ||
            valor === "" ||
            Number.isNaN(Number(valor))
        ) {
            return "-";
        }

        return Number(valor)
            .toFixed(2)
            .replace(".", ",") + "%";
    }


    function classeEvolucao(valor) {

        if (
            valor === null ||
            valor === undefined ||
            valor === "-" ||
            Number.isNaN(Number(valor))
        ) {
            return "vazio";
        }

        return Number(valor) > 0
            ? "positivo"
            : Number(valor) < 0
                ? "negativo"
                : "vazio";
    }


    function textoEvolucao(valor, formatado) {

        if (
            formatado !== undefined &&
            formatado !== null &&
            formatado !== ""
        ) {
            return formatado;
        }

        if (
            valor === null ||
            valor === undefined ||
            Number.isNaN(Number(valor))
        ) {
            return "-";
        }

        const n = Number(valor);

        const sinal =
            n > 0 ? "+" : "";

        return (
            sinal +
            n.toFixed(2).replace(".", ",") +
            "%"
        );
    }


    function valorNumerico(item, campo) {

        const valor =
            Number(item?.[campo]);

        if (Number.isNaN(valor)) {
            return null;
        }

        return valor;
    }


    function calcularEvolucao(
        lista,
        campo,
        indice
    ) {

        const atual =
            valorNumerico(
                lista[indice],
                campo
            );


        if (atual === null) {
            return null;
        }


        if (indice === 0) {
            return null;
        }


        const anterior =
            valorNumerico(
                lista[indice - 1],
                campo
            );


        if (
            anterior === null ||
            anterior === 0
        ) {
            return null;
        }


        return (
            (atual - anterior) /
            anterior
        ) * 100;
    }


    function calcularEvolucaoAcumulada(
        lista,
        campo,
        indice
    ) {

        if (indice === 0) {
            return null;
        }

        const atual =
            valorNumerico(
                lista[indice],
                campo
            );

        const base =
            valorNumerico(
                lista[0],
                campo
            );


        if (
            atual === null ||
            base === null ||
            base === 0
        ) {
            return null;
        }


        return (
            (atual - base) /
            base
        ) * 100;
    }


    function classePercentual(valor) {

        const numero =
            Number(valor);


        if (Number.isNaN(numero)) {
            return "";
        }


        if (numero <= 48.6) {
            return "percentual-verde";
        }


        if (numero <= 51.3) {
            return "percentual-alerta";
        }


        if (numero <= 54) {
            return "percentual-prudencial";
        }


        return "percentual-maximo";
    }


    function percentualHTML(item) {

        const valor =
            item.percentual;


        if (
            valor === null ||
            valor === undefined ||
            valor === "" ||
            Number.isNaN(Number(valor))
        ) {

            return `
                <span class="vazio">
                    -
                </span>
            `;
        }


        const texto =
            item.percentualFormatado ||
            formatarPercentual(valor);


        return `
            <span
                class="badge-percentual ${classePercentual(valor)}"
            >
                ${escaparHTML(texto)}
            </span>
        `;
    }


    async function buscarJSON(url) {

        const resposta =
            await fetch(url, {
                method: "GET",
                cache: "no-store"
            });


        const texto =
            await resposta.text();


        let dados;


        try {

            dados =
                JSON.parse(texto);

        } catch {

            throw new Error(
                "O servidor não retornou JSON válido."
            );
        }


        if (
            !resposta.ok ||
            dados.sucesso === false
        ) {

            throw new Error(
                dados.erro ||
                "Erro na comunicação com o servidor."
            );
        }


        return dados;
    }


    async function carregarInicial() {

        try {

            definirCarregando(
                "Conectando ao TCE-PR..."
            );


            const statusDados =
                await buscarJSON(
                    "/api/status"
                );


            const municipiosDados =
                await buscarJSON(
                    "/api/municipios"
                );


            municipioSelect.innerHTML =
                '<option value="">Selecione o município</option>';


            for (
                const municipio
                of municipiosDados.municipios || []
            ) {

                const option =
                    document.createElement(
                        "option"
                    );

                option.value =
                    municipio.id;

                option.textContent =
                    municipio.nome;

                municipioSelect.appendChild(
                    option
                );
            }


            anoSelect.innerHTML =
                '<option value="">Selecione o ano</option>';


            const anos =
                statusDados.anos || [];


            for (const ano of anos) {

                const option =
                    document.createElement(
                        "option"
                    );

                option.value =
                    ano;

                option.textContent =
                    ano;

                anoSelect.appendChild(
                    option
                );
            }


            if (anos.length) {

                anoSelect.value =
                    anos.includes("2025")
                        ? "2025"
                        : anos[anos.length - 1];
            }


            municipioSelect.disabled = false;

            anoSelect.disabled = false;

            btnConsultar.disabled = false;

            btnHistorico.disabled = false;


            definirSucesso(
                `TCE-PR conectado • ${municipiosDados.municipios.length} municípios`
            );


            mensagem.textContent =
                "Selecione o município e o ano para iniciar a consulta.";

        } catch (erro) {

            console.error(erro);

            definirErroStatus(
                "Erro ao conectar ao TCE-PR"
            );

            mostrarErro(
                erro.message
            );

            mensagem.textContent =
                "Não foi possível carregar os dados iniciais.";
        }
    }


    function validarSelecao() {

        const municipio =
            municipioSelect.value;

        const ano =
            anoSelect.value;


        if (!municipio) {
            throw new Error(
                "Selecione um município."
            );
        }


        if (!ano) {
            throw new Error(
                "Selecione um ano."
            );
        }


        return {
            municipio,
            ano
        };
    }


    function limparResultados() {

        resultadoUnico.classList.remove(
            "visivel"
        );

        resultadoHistorico.classList.remove(
            "visivel"
        );

        esconderErro();
    }


    function preencherConsultaUnica(dados) {

        const item =
            dados.dados || dados;


        document.getElementById(
            "badgeAno"
        ).textContent =
            `Ano ${item.ano || anoSelect.value}`;


        document.getElementById(
            "rcl"
        ).textContent =
            item.rclFormatada ||
            formatarMoeda(item.rcl);


        document.getElementById(
            "rclAjustada"
        ).textContent =
            item.rclAjustadaFormatada ||
            formatarMoeda(
                item.rclAjustada
            );


        document.getElementById(
            "dtp"
        ).textContent =
            item.dtpFormatado ||
            formatarMoeda(item.dtp);


        document.getElementById(
            "percentual"
        ).textContent =
            item.percentualFormatado ||
            formatarPercentual(
                item.percentual
            );


        document.getElementById(
            "limiteAlerta"
        ).textContent =
            item.limiteAlertaFormatado ||
            formatarMoeda(
                item.limiteAlerta
            );


        document.getElementById(
            "limitePrudencial"
        ).textContent =
            item.limitePrudencialFormatado ||
            formatarMoeda(
                item.limitePrudencial
            );


        document.getElementById(
            "limiteMaximo"
        ).textContent =
            item.limiteMaximoFormatado ||
            formatarMoeda(
                item.limiteMaximo
            );


        document.getElementById(
            "fundeb"
        ).textContent =
            item.fundebFormatado ||
            formatarMoeda(
                item.fundeb
            );


        const situacao =
            document.getElementById(
                "situacao"
            );


        situacao.textContent =
            item.situacao || "-";


        situacao.className =
            "situacao";


        const situacaoTexto =
            String(
                item.situacao || ""
            ).toUpperCase();


        if (
            situacaoTexto.includes("ACIMA") ||
            situacaoTexto.includes("MÁXIMO") ||
            situacaoTexto.includes("MAXIMO")
        ) {

            situacao.classList.add(
                "perigo"
            );

        } else if (
            situacaoTexto.includes("PRUDENCIAL") ||
            situacaoTexto.includes("ALERTA")
        ) {

            situacao.classList.add(
                "alerta"
            );
        }


        const origem =
            document.getElementById(
                "fundebOrigem"
            );


        if (item.fundebOrigem) {

            origem.textContent =
                "Fonte: " +
                item.fundebOrigem;

            origem.style.display =
                "block";

        } else {

            origem.style.display =
                "none";
        }


        resultadoUnico.classList.add(
            "visivel"
        );
    }


    function preencherHistorico(dados) {

        const listaCompleta =
            dados.dados || [];

        const lista =
            listaCompleta.filter(item => {
                const ano = Number(item.ano);
                return ano >= 2019 && ano <= 2025;
            });


        dadosHistoricoAtual =
            lista;


        historicoTabela.innerHTML =
            "";

        const tabelaWrapper =
            document.querySelector("#historicoTabela .tabela-wrapper");

        if (tabelaWrapper) {
            tabelaWrapper.scrollLeft = 0;
        }


        document.getElementById(
            "badgeMunicipio"
        ).textContent =
            municipioSelect
                .options[
                    municipioSelect.selectedIndex
                ]
                ?.textContent || "";


        if (!lista.length) {

            historicoTabela.innerHTML = `
                <tr>
                    <td
                        colspan="9"
                        class="vazio"
                        style="text-align:center;padding:30px;"
                    >
                        Nenhum dado encontrado.
                    </td>
                </tr>
            `;

            resultadoHistorico.classList.add(
                "visivel"
            );

            return;
        }


        function linhaIndicador(
            nome,
            valores
        ) {

            const tr =
                document.createElement(
                    "tr"
                );


            tr.className =
                "linha-indicador";


            tr.innerHTML = `
                <td>${nome}</td>
                <td class="tipo">Valor</td>
                ${valores}
            `;


            historicoTabela.appendChild(
                tr
            );
        }


        function linhaSub(
            nome,
            valores
        ) {

            const tr =
                document.createElement(
                    "tr"
                );


            tr.className =
                "linha-sub";


            tr.innerHTML = `
                <td
                    style="
                        padding-left:24px;
                        color:#68788b;
                    "
                >
                    ${nome}
                </td>

                <td class="tipo">%</td>

                ${valores}
            `;


            historicoTabela.appendChild(
                tr
            );
        }


        const rclAjustadaValores =
            lista.map(item => `

                <td class="valor-forte">

                    ${escaparHTML(
                        item.rclAjustadaFormatada ||
                        formatarMoeda(
                            item.rclAjustada
                        )
                    )}

                </td>

            `).join("");


        linhaIndicador(
            "RREO — Receita Corrente Líquida Ajustada",
            rclAjustadaValores
        );


        const evolucaoRCLAjustada =
            lista.map(
                (item, indice) => {

                    let valor =
                        item.evolucaoRCLAjustada;

                    let formatado =
                        item.evolucaoRCLAjustadaFormatada;


                    if (
                        valor === undefined ||
                        valor === null
                    ) {

                        valor =
                            calcularEvolucao(
                                listaCompleta,
                                "rclAjustada",
                                indice
                            );

                        formatado =
                            undefined;
                    }


                    return `

                        <td
                            class="${classeEvolucao(valor)}"
                        >

                            ${escaparHTML(
                                textoEvolucao(
                                    valor,
                                    formatado
                                )
                            )}

                        </td>

                    `;
                }
            ).join("");


        linhaSub(
            "Evolução anual",
            evolucaoRCLAjustada
        );


        const acumuladaRCLAjustada =
            lista.map(
                (item, indice) => {

                    let valor =
                        item.evolucaoRCLAjustadaAcumulada;

                    let formatado =
                        item.evolucaoRCLAjustadaAcumuladaFormatada;


                    if (
                        valor === undefined ||
                        valor === null
                    ) {

                        valor =
                            calcularEvolucaoAcumulada(
                                listaCompleta,
                                "rclAjustada",
                                indice
                            );

                        formatado =
                            undefined;
                    }


                    return `

                        <td
                            class="${classeEvolucao(valor)}"
                        >

                            ${escaparHTML(
                                textoEvolucao(
                                    valor,
                                    formatado
                                )
                            )}

                        </td>

                    `;
                }
            ).join("");


        linhaSub(
            "Evolução acumulada",
            acumuladaRCLAjustada
        );


        const dtpValores =
            lista.map(item => `

                <td class="valor-forte">

                    ${escaparHTML(
                        item.dtpFormatado ||
                        formatarMoeda(
                            item.dtp
                        )
                    )}

                </td>

            `).join("");


        linhaIndicador(
            "RGF — Despesa Total com Pessoal",
            dtpValores
        );


        const evolucaoDTP =
            lista.map(
                (item, indice) => {

                    let valor =
                        item.evolucaoDTP;

                    let formatado =
                        item.evolucaoDTPFormatada;


                    if (
                        valor === undefined ||
                        valor === null
                    ) {

                        valor =
                            calcularEvolucao(
                                listaCompleta,
                                "dtp",
                                indice
                            );

                        formatado =
                            undefined;
                    }


                    return `

                        <td
                            class="${classeEvolucao(valor)}"
                        >

                            ${escaparHTML(
                                textoEvolucao(
                                    valor,
                                    formatado
                                )
                            )}

                        </td>

                    `;
                }
            ).join("");


        linhaSub(
            "Evolução anual",
            evolucaoDTP
        );


        const acumuladaDTP =
            lista.map(
                (item, indice) => {

                    let valor =
                        item.evolucaoDTPAcumulada;

                    let formatado =
                        item.evolucaoDTPAcumuladaFormatada;


                    if (
                        valor === undefined ||
                        valor === null
                    ) {

                        valor =
                            calcularEvolucaoAcumulada(
                                listaCompleta,
                                "dtp",
                                indice
                            );

                        formatado =
                            undefined;
                    }


                    return `

                        <td
                            class="${classeEvolucao(valor)}"
                        >

                            ${escaparHTML(
                                textoEvolucao(
                                    valor,
                                    formatado
                                )
                            )}

                        </td>

                    `;
                }
            ).join("");


        linhaSub(
            "Evolução acumulada",
            acumuladaDTP
        );


        const percentualValores =
            lista.map(item => `

                <td>
                    ${percentualHTML(item)}
                </td>

            `).join("");


        linhaIndicador(
            "RGF — Despesa Total com Pessoal %",
            percentualValores
        );


        const fundebValores =
            lista.map(item => `

                <td class="fundeb-col">

                    ${escaparHTML(
                        item.fundebFormatado ||
                        formatarMoeda(
                            item.fundeb
                        )
                    )}

                </td>

            `).join("");


        linhaIndicador(
            "RREO — Transferências FUNDEB",
            fundebValores
        );


        const evolucaoFundeb =
            lista.map(
                (item, indice) => {

                    let valor =
                        item.evolucaoFundeb;

                    let formatado =
                        item.evolucaoFundebFormatada;


                    if (
                        valor === undefined ||
                        valor === null
                    ) {

                        valor =
                            calcularEvolucao(
                                listaCompleta,
                                "fundeb",
                                indice
                            );

                        formatado =
                            undefined;
                    }


                    return `

                        <td
                            class="${classeEvolucao(valor)}"
                        >

                            ${escaparHTML(
                                textoEvolucao(
                                    valor,
                                    formatado
                                )
                            )}

                        </td>

                    `;
                }
            ).join("");


        linhaSub(
            "Evolução anual",
            evolucaoFundeb
        );


        const acumuladaFundeb =
            lista.map(
                (item, indice) => {

                    let valor =
                        item.evolucaoFundebAcumulada;

                    let formatado =
                        item.evolucaoFundebAcumuladaFormatada;


                    if (
                        valor === undefined ||
                        valor === null
                    ) {

                        valor =
                            calcularEvolucaoAcumulada(
                                listaCompleta,
                                "fundeb",
                                indice
                            );

                        formatado =
                            undefined;
                    }


                    return `

                        <td
                            class="${classeEvolucao(valor)}"
                        >

                            ${escaparHTML(
                                textoEvolucao(
                                    valor,
                                    formatado
                                )
                            )}

                        </td>

                    `;
                }
            ).join("");


        linhaSub(
            "Evolução acumulada",
            acumuladaFundeb
        );


        resultadoHistorico.classList.add(
            "visivel"
        );


        mostrarTabelaHistorico();
    }


    function mostrarTabelaHistorico() {

        historicoTabelaContainer.style.display =
            "block";

        historicoGrafico.classList.remove(
            "visivel"
        );

        btnAlternarHistorico.textContent =
            "Ver gráfico";
    }


    function mostrarGraficoHistorico() {

        historicoTabelaContainer.style.display =
            "none";

        historicoGrafico.classList.add(
            "visivel"
        );

        btnAlternarHistorico.textContent =
            "Voltar ao histórico";


        criarGraficoHistorico();
    }


    function criarGraficoHistorico() {

        if (
            !dadosHistoricoAtual.length
        ) {
            return;
        }


        const canvas =
            document.getElementById(
                "graficoHistoricoCanvas"
            );


        if (graficoHistorico) {

            graficoHistorico.destroy();
        }


        const anos =
            dadosHistoricoAtual.map(
                item =>
                    String(item.ano)
            );


        const formatadorEixo =
            new Intl.NumberFormat(
                "pt-BR",
                {
                    maximumFractionDigits: 0
                }
            );


        graficoHistorico =
            new Chart(
                canvas,
                {

                    type: "bar",

                    data: {

                        labels: anos,

                        datasets: [

                            {
                                label:
                                    "RCL Ajustada",

                                data:
                                    dadosHistoricoAtual.map(
                                        item =>
                                            Number(
                                                item.rclAjustada
                                            ) || 0
                                    ),

                                backgroundColor:
                                    "#2477a8",

                                borderRadius: 0,

                                borderSkipped: false
                            },


                            {
                                label:
                                    "FUNDEB",

                                data:
                                    dadosHistoricoAtual.map(
                                        item =>
                                            Number(
                                                item.fundeb
                                            ) || 0
                                    ),

                                backgroundColor:
                                    "#30964f",

                                borderRadius: 0,

                                borderSkipped: false
                            },


                            {
                                label:
                                    "DTP",

                                data:
                                    dadosHistoricoAtual.map(
                                        item =>
                                            Number(
                                                item.dtp
                                            ) || 0
                                    ),

                                backgroundColor:
                                    "#d44949",

                                borderRadius: 0,

                                borderSkipped: false
                            }

                        ]
                    },


                    options: {

                        responsive: true,

                        maintainAspectRatio: false,

                        interaction: {
                            mode: "index",
                            intersect: false
                        },


                        plugins: {

                            legend: {
                                position: "top",

                                labels: {
                                    usePointStyle: false,
                                    boxWidth: 34,
                                    boxHeight: 10,
                                    padding: 15,

                                    font: {
                                        size: 12
                                    }
                                }
                            },


                            tooltip: {

                                backgroundColor:
                                    "#18232f",

                                padding: 12,

                                callbacks: {

                                    label:
                                        function(context) {

                                            return (
                                                " " +
                                                context.dataset.label +
                                                ": " +
                                                Number(
                                                    context.raw
                                                ).toLocaleString(
                                                    "pt-BR",
                                                    {
                                                        minimumFractionDigits: 2,
                                                        maximumFractionDigits: 2
                                                    }
                                                )
                                            );
                                        }
                                }
                            }
                        },


                        scales: {

                            x: {

                                grid: {
                                    color:
                                        "rgba(100,120,140,.14)"
                                },

                                ticks: {
                                    color:
                                        "#617286",

                                    font: {
                                        size: 12
                                    }
                                }
                            },


                            y: {

                                beginAtZero: true,

                                grid: {
                                    color:
                                        "rgba(100,120,140,.14)"
                                },

                                ticks: {

                                    color:
                                        "#617286",

                                    font: {
                                        size: 11
                                    },

                                    callback:
                                        function(value) {

                                            return formatadorEixo.format(
                                                value
                                            );
                                        }
                                }
                            }
                        }
                    }
                }
            );
    }


    async function exportarExcel() {

        if (
            !dadosHistoricoAtual.length
        ) {
            return;
        }

        const municipio =
            municipioSelect.value;

        if (!municipio) {
            return;
        }

        const resposta =
            await fetch(
                `/api/exportar-excel?municipio=${encodeURIComponent(municipio)}`,
                {
                    cache: "no-store"
                }
            );

        if (!resposta.ok) {
            let erro = "Não foi possível gerar o Excel.";

            try {
                const dados =
                    await resposta.json();

                erro =
                    dados.erro ||
                    erro;
            } catch {}

            throw new Error(erro);
        }

        const arquivo =
            await resposta.blob();

        const url =
            URL.createObjectURL(
                arquivo
            );

        const link =
            document.createElement(
                "a"
            );

        link.href = url;

        const municipioNome =
            municipioSelect
                .options[
                    municipioSelect.selectedIndex
                ]
                ?.textContent ||
                "Municipio";

        link.download =
            `Historico_${municipioNome
                .normalize("NFD")
                .replace(/[\u0300-\u036f]/g, "")
                .replace(/[^a-zA-Z0-9 ]/g, "")
                .replace(/\s+/g, "_")}_2019-2025.xlsx`;

        document.body.appendChild(
            link
        );

        link.click();

        link.remove();

        URL.revokeObjectURL(
            url
        );
    }


    async function consultar() {

        try {

            limparResultados();

            const {
                municipio,
                ano
            } = validarSelecao();


            bloquearBotoes();

            iniciarLoading(
                "unica"
            );


            definirCarregando(
                "Consultando TCE-PR..."
            );


            mensagem.textContent =
                "Buscando os dados do município. Aguarde...";


            const resposta =
                await fetch(
                    "/api/consulta",
                    {

                        method: "POST",

                        headers: {
                            "Content-Type":
                                "application/json"
                        },

                        body:
                            JSON.stringify({
                                municipio,
                                ano
                            })
                    }
                );


            const textoResposta =
                await resposta.text();


            let dados;


            try {

                dados =
                    JSON.parse(
                        textoResposta
                    );

            } catch {

                throw new Error(
                    "O servidor não retornou JSON válido."
                );
            }


            if (
                !resposta.ok ||
                dados.sucesso === false
            ) {

                throw new Error(
                    dados.erro ||
                    "Não foi possível realizar a consulta."
                );
            }


            preencherConsultaUnica(
                dados
            );


            definirSucesso(
                "Consulta concluída"
            );


            mensagem.textContent =
                "Dados encontrados com sucesso.";


            window.scrollTo({
                top:
                    resultadoUnico.offsetTop - 30,
                behavior:
                    "smooth"
            });


        } catch (erro) {

            console.error(erro);

            definirErroStatus(
                "Erro na consulta"
            );

            mostrarErro(
                erro.message
            );

            mensagem.textContent =
                "Não foi possível concluir a consulta.";

        } finally {

            finalizarLoading();

            desbloquearBotoes();
        }
    }


    async function consultarHistorico() {

        try {

            limparResultados();

            const {
                municipio
            } = validarSelecao();


            bloquearBotoes();

            iniciarLoading(
                "historico"
            );


            definirCarregando(
                "Consultando histórico no TCE-PR..."
            );


            mensagem.textContent =
                "Consultando os anos de 2019 a 2025. Aguarde...";


            const url =
                `/historico?municipio=${encodeURIComponent(municipio)}` +
                `&anoInicial=2019` +
                `&anoFinal=2025` +
                `&periodo=12`;


            const dados =
                await buscarJSON(
                    url
                );


            preencherHistorico(
                dados
            );


            definirSucesso(
                "Consulta histórica concluída"
            );


            mensagem.textContent =
                "Histórico de 2019 a 2025 carregado.";


            window.scrollTo({
                top:
                    resultadoHistorico.offsetTop - 30,
                behavior:
                    "smooth"
            });


        } catch (erro) {

            console.error(erro);

            definirErroStatus(
                "Erro na consulta histórica"
            );

            mostrarErro(
                erro.message
            );

            mensagem.textContent =
                "Não foi possível concluir a consulta histórica.";

        } finally {

            finalizarLoading();

            desbloquearBotoes();
        }
    }


    btnAlternarHistorico.addEventListener(
        "click",
        function() {

            const graficoVisivel =
                historicoGrafico.classList.contains(
                    "visivel"
                );


            if (graficoVisivel) {

                mostrarTabelaHistorico();

            } else {

                mostrarGraficoHistorico();
            }
        }
    );


    btnExportarExcel.addEventListener(
        "click",
        exportarExcel
    );


    btnConsultar.addEventListener(
        "click",
        consultar
    );


    btnHistorico.addEventListener(
        "click",
        consultarHistorico
    );


    carregarInicial();
