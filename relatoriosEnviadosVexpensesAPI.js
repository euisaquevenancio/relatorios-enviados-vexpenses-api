/*
    @euisaquevenancio - 04/10/2026
    Automação para captura de relatórios enviados, pendentes, que estão aguardando a mais de 10 dias no Vexpenses das mantenedoras ABEC, SOME, UBEE e UNBEC, desenvolvido com Node.js.
    
    Instalando todas bibliotecas de uma vez via terminal:
    npm install

    Instalando as bibliotecas manualmente via terminal:
    npm install dotenv
    npm install exceljs

    Executando o código via terminal:
    node relatoriosEnviadosVexpensesAPI.js
*/

require("dotenv").config(); // Manipulação de variáveis de ambiente .env
const FS = require("fs"); // Manipulação de arquivos - File System
const EXCEL_JS = require("exceljs"); // Manipulação de arquivos Excel
const PATH = require("path"); // Manipulação de caminhos de arquivos
const { exec } = require("child_process"); // Execução de comandos do sistema operacional - necessário para abrir o arquivo Excel no final do processo

const TOKEN_VEXPENSES = process.env.TOKEN_VEXPENSES;
const HOJE = new Date();
let relatorios = [];
let contadorRelatorios = 0;

async function main() {
    const HORARIO_INICIO = new Date().toLocaleTimeString("pt-BR");
    console.log(`\n🤖 Início da execução do script às ${HORARIO_INICIO}.`);

    let dataRelatoriosEnviados = await capturarRelatoriosAprovados();

    // Percorrendo todos os relatórios aprovados
    for (let i = 0; i < dataRelatoriosEnviados.length; i++) {
        // Capturando as informações das despesas e solicitante
        let dataDespesas = await capturarDespesasRelatorio(dataRelatoriosEnviados[i].id);

        let dataCriacaoRelatorio = dataRelatoriosEnviados[i].created_at.substring(0, 10).replaceAll("-", "/").split("/").reverse().join("/");
        let dataAtualizacao = dataRelatoriosEnviados[i].updated_at.substring(0, 10).replaceAll("-", "/").split("/").reverse().join("/");

        const DIFERENCA_HOJE_ULTIMA_ATUALIZACAO = await gerarDiferencaDatas(dataAtualizacao);
        if (DIFERENCA_HOJE_ULTIMA_ATUALIZACAO < 11) continue;

        contadorRelatorios++;
        if (contadorRelatorios == 1) console.log();

        let CENTRO_CUSTOS = "VERIFICAR";
        if (dataDespesas.data.expenses.data[0].costs_center.data.integration_id != null) CENTRO_CUSTOS = dataDespesas.data.expenses.data[0].costs_center.data.integration_id;

        const MOEDA = dataDespesas.data.expenses.data[0].original_currency_iso;
        const NOME_SOLICITANTE = dataDespesas.data.user.data.name;
        let tipoRelatorio = "VERIFICAR";
        if (dataDespesas.data.payment_method.data.description != undefined) {
            tipoRelatorio = dataDespesas.data.payment_method.data.description.toUpperCase();
        }

        if (tipoRelatorio == "REEMBOLSO GASTOS EVENTUAIS") {
            tipoRelatorio = "REEMBOLSO";
        } else if (tipoRelatorio == "PRESTAÇÃO DE CONTAS DO ADIANTAMENTO" || tipoRelatorio == "PRESTAÇÃO DE CONTAS ADIANTAMENTO - MOEDA ESTRANGEIRA") {
            tipoRelatorio = "PRESTAÇÃO DE CONTAS";
        }

        dataDespesas = dataDespesas.data.expenses.data;

        const QUANTIDADE_DESPESAS = dataDespesas.length;
        let valorTotalDespesas = 0;

        for (let j = 0; j < QUANTIDADE_DESPESAS; j++) {
            valorTotalDespesas += dataDespesas[j].value;
        }

        let mantenedora = parseInt(dataDespesas[0].apportionment.data[0].integration_id[0]);
        if (mantenedora == 1) {
            mantenedora = "1 ABEC";
        } else if (mantenedora == 2) {
            mantenedora = "2 SOME";
        } else if (mantenedora == 3) {
            mantenedora = "3 UBEE";
        } else {
            mantenedora = "4 UNBEC";
        }

        if (MOEDA == "BRL") {
            valorTotalDespesas = valorTotalDespesas.toFixed(2).replace(".", ",");
        } else if (MOEDA == "USD") {
            valorTotalDespesas = "$" + valorTotalDespesas.toFixed(2);
        } else {
            valorTotalDespesas = "€" + valorTotalDespesas.toFixed(2);
        }

        // Adicionando o relatório na lista
        relatorios.push({
            numeroRelatorio: dataRelatoriosEnviados[i].id,
            mantenedora: mantenedora,
            nomeSolicitante: NOME_SOLICITANTE,
            moedaRelatorio: MOEDA,
            valorRelatorio: valorTotalDespesas,
            tipoRelatorio: tipoRelatorio,
            dataCriacaoRelatorio: dataCriacaoRelatorio,
            dataAtualizacao: dataAtualizacao,
            diferencaDatasHojeEUltimaAtualizacao: DIFERENCA_HOJE_ULTIMA_ATUALIZACAO,
            centroCustos: CENTRO_CUSTOS
        });
    }

    if (contadorRelatorios > 0) {
        console.log(`✅ Novos relatórios aprovados: ${contadorRelatorios}\n`);
        for (let i = 0; i < relatorios.length; i++) {
            console.log(`✅ #${(i+1)} | ${relatorios[i].numeroRelatorio} | ${relatorios[i].mantenedora} | ${relatorios[i].nomeSolicitante} | ${relatorios[i].valorRelatorio} | ${relatorios[i].dataCriacaoRelatorio} | ${relatorios[i].dataAtualizacao} | ${relatorios[i].tipoRelatorio} | ${relatorios[i].centroCustos}`);
        }
    } else {
        console.log(`\n✅ Não há novos relatórios aprovados: ${contadorRelatorios}`);
    }

    const HORARIO_FIM = new Date().toLocaleTimeString("pt-BR");
    console.log(`\n🤖 Fim da execução do script às ${HORARIO_FIM}.`);
    console.log(`🕓 Tempo de execução: ${calcularDiferencaHoras(HORARIO_INICIO, HORARIO_FIM)}.\n`);

    await salvarRelatorios();
}

// Função para capturar os relatórios aprovados
async function capturarRelatoriosAprovados() {
    const responseRelatoriosAprovados = await fetch(
        "https://api.vexpenses.com/v2/reports/status/ENVIADO",
        {
            method: "GET",
            headers: {
                "Authorization": TOKEN_VEXPENSES,
                "Content-Type": "application/json"
            }
        }
    );

    let dataRelatoriosEnviados = await responseRelatoriosAprovados.json();
    return dataRelatoriosEnviados.data;
}

// Função para capturar as despesas do relatório informado
async function capturarDespesasRelatorio(idRelatorio) {
    const responseDespesa = await fetch(
        `https://api.vexpenses.com/v2/reports/${idRelatorio}?include=expenses%2Cexpenses.apportionment%2Cexpenses.expense_type%2Cexpenses.fueling%2Cuser%2Chistory%2Capproval%2Cpayment_method%2Cexpenses.costs_center`,
        {
            method: "GET",
            headers: {
                "Authorization": TOKEN_VEXPENSES,
                "Content-Type": "application/json"
            }
        }
    );

    let dataDespesas = await responseDespesa.json();
    return dataDespesas;
}

// Função para calcular a diferença entre dois horários no formato HH:mm:ss
function calcularDiferencaHoras(HORARIO_INICIO, HORARIO_FIM) {
    try {
        // Quebra as strings em partes
        const [h1, m1, s1] = HORARIO_INICIO.split(":").map(Number);
        const [h2, m2, s2] = HORARIO_FIM.split(":").map(Number);

        // Cria objetos Date no mesmo dia
        const dataBase = new Date();
        const date1 = new Date(dataBase.getFullYear(), dataBase.getMonth(), dataBase.getDate(), h1, m1, s1 || 0);
        const date2 = new Date(dataBase.getFullYear(), dataBase.getMonth(), dataBase.getDate(), h2, m2, s2 || 0);

        // Calcula a diferença em milissegundos
        let diffMs = date2 - date1;

        // Se negativo, inverte
        const negativo = diffMs < 0;
        diffMs = Math.abs(diffMs);

        // Converte para horas, minutos e segundos
        const horas = Math.floor(diffMs / (1000 * 60 * 60));
        const minutos = Math.floor((diffMs % (1000 * 60 * 60)) / (1000 * 60));
        const segundos = Math.floor((diffMs % (1000 * 60)) / 1000);

        return `${negativo ? '-' : ''}${String(horas).padStart(2, '0')}:${String(minutos).padStart(2, '0')}:${String(segundos).padStart(2, '0')}`;
    } catch (err) {
        console.error("Erro ao calcular diferença de horas: ", err);
        return null;
    }
}

// Função para converter um valor para o tipo data
function converterParaData(data) {
    if (!data) return null;

    // Se já for um objeto Date
    if (data instanceof Date) {
        return data;
    }

    // Se vier como string no formato dd/mm/aaaa
    if (typeof data === "string") {
        const partes = data.trim().split("/");

        if (partes.length === 3) {
            const dia = Number(partes[0]);
            const mes = Number(partes[1]);
            const ano = Number(partes[2]);

            return new Date(ano, mes - 1, dia);
        }
    }

    return null;
}

// Função para salvar os relatórios em Excel
async function salvarRelatorios() {
    const PASTA_DESTINO = PATH.join(__dirname, "dados");

    // Garante que a pasta exista
    if (!FS.existsSync(PASTA_DESTINO)) {
        FS.mkdirSync(PASTA_DESTINO, { recursive: true });
    }

    // Captura o arquivo Excel
    const ARQUIVO_EXCEL = PATH.join(PASTA_DESTINO, "relatoriosEnviados.xlsx");

    const WORKBOOK = new EXCEL_JS.Workbook();
    // Se o arquivo existe, lê
    if (FS.existsSync(ARQUIVO_EXCEL)) {
        await WORKBOOK.xlsx.readFile(ARQUIVO_EXCEL);
    }

    let planilha = WORKBOOK.getWorksheet("Relatorios");
    // Se a planilha não estiver estruturada, adiciona o cabeçalho
    if (!planilha) {
        planilha = WORKBOOK.addWorksheet("Relatorios");

        planilha.addRow([
            "N° RELATÓRIO (VEXPESES)",
            "MANTENEDORA",
            "NOME",
            "VALOR",
            "DATA CRIAÇÃO",
            "DATA ATUALIZAÇÃO",
            "DIFERENÇA HOJE E ULTIMA ATUALIZAÇÃO",
            "TIPO DE DESPESA",
            "CR (SOMENTE PARA UBEE E UNBEC)",
            "OBSERVAÇÃO"
        ]);
    }

    // Aplica a largura nas colunas
    planilha.columns = [
        { key: "N° RELATÓRIO (VEXPESES)", width: 8 },
        { key: "MANTENEDORA", width: 8 },
        { key: "NOME", width: 8 },
        { key: "VALOR", width: 8 },
        { key: "DATA CRIAÇÃO", width: 8 },
        { key: "DATA ATUALIZAÇÃO", width: 8 },
        { key: "DIFERENÇA HOJE E ULTIMA ATUALIZAÇÃO", width: 8 },
        { key: "TIPO DE DESPESA", width: 8 },
        { key: "CR (SOMENTE PARA UBEE E UNBEC)", width: 8 },
        { key: "OBSERVAÇÃO", width: 8 }
    ];

    // Adicionando os tickets na planilha
    if (relatorios.length > 0) {
        for (let i = 0; i < relatorios.length; i++) {
            const DATA_CRIACAO_RELATORIO = converterParaData(relatorios[i].dataCriacaoRelatorio);
            const DATA_ATUALIZACAO_RELATORIO = converterParaData(relatorios[i].dataAtualizacao);

            let observacao = "";

            if (relatorios[i].valorRelatorio.includes(".")) {
                observacao = "MOEDA ESTRANGEIRA | " + relatorios[i].valorRelatorio + " = R$ ";
            }

            const LINHA = planilha.addRow([
                relatorios[i].numeroRelatorio,
                relatorios[i].mantenedora,
                relatorios[i].nomeSolicitante,
                relatorios[i].valorRelatorio,
                DATA_CRIACAO_RELATORIO,
                DATA_ATUALIZACAO_RELATORIO,
                relatorios[i].diferencaDatasHojeEUltimaAtualizacao,
                relatorios[i].tipoRelatorio,
                relatorios[i].centroCustos,
                observacao
            ]);

            if (relatorios[i].valorRelatorio.includes(".")) {
                // Linha inteira em negrito
                LINHA.font = { bold: true };

                // Coluna VALOR em verde
                LINHA.getCell(4).fill = {
                    type: "pattern",
                    pattern: "solid",
                    fgColor: {
                        argb: "C1F0C8"
                    }
                };

                // Coluna OBSERVAÇÃO em verde
                LINHA.getCell(10).fill = {
                    type: "pattern",
                    pattern: "solid",
                    fgColor: {
                        argb: "C1F0C8"
                    }
                };
            }
        }
    }

    if (planilha.rowCount > 1) {
        // Remove os relatórios da tabela antiga
        if (planilha.model.tables) {
            planilha.model.tables = [];
        }

        const LINHAS_VALIDAS = planilha.getSheetValues().slice(2) // Remove o cabeçalho
                                                       .filter(linha => Array.isArray(linha)) // Remove undefined
                                                       .map(linha => linha.slice(1)); // Remove índice fantasma
        
        planilha.addTable({
            name: "TabelaRelatorios",
            ref: "A1",
            headerRow: true,
            style: {
                theme: "TableStyleLight1"
            },
            columns: [
                { name: "N° RELATÓRIO (VEXPESES)" },
                { name: "MANTENEDORA" },
                { name: "NOME" },
                { name: "VALOR" },
                { name: "DATA CRIAÇÃO" },
                { name: "DATA ATUALIZAÇÃO" },
                { name: "DIFERENÇA HOJE E ULTIMA ATUALIZAÇÃO"},
                { name: "TIPO DE DESPESA" },
                { name: "CR (SOMENTE PARA UBEE E UNBEC)" },
                { name: "OBSERVAÇÃO" }
            ],
            rows: LINHAS_VALIDAS
        });
    }

    if (relatorios.length > 0) {
        // Salva o arquivo excel
        await WORKBOOK.xlsx.writeFile(ARQUIVO_EXCEL);

        // Adicionando os novos relatórios que devem ser ignorados (já estão na planilha)
        const ARQUIVO_RELATORIOS_IGNORADOS = PATH.join(PASTA_DESTINO, "relatoriosIgnorados.txt");

        let conteudoExistenteRelatoriosIgnorados = "";
        // Se o arquivo existir, lê o conteúdo
        if (FS.existsSync(ARQUIVO_RELATORIOS_IGNORADOS)) {
            conteudoExistenteRelatoriosIgnorados = FS.readFileSync(ARQUIVO_RELATORIOS_IGNORADOS, "utf-8");
        }

        let novoConteudoRelatoriosIgnorados = "";
        for (const relatorioAtual of relatorios) {
            novoConteudoRelatoriosIgnorados += `\n${relatorioAtual.numeroRelatorio}`;
        }

        // Só escreve no TXT se tiver conteúdo novo
        if (novoConteudoRelatoriosIgnorados != "") {
            FS.appendFileSync(ARQUIVO_RELATORIOS_IGNORADOS, novoConteudoRelatoriosIgnorados, "utf-8");
        }

        exec(`start "" "${ARQUIVO_EXCEL}"`);
    }
}

async function gerarDiferencaDatas(dataAtualizacao) {
    const [dia, mes, ano] = dataAtualizacao.split("/");
    dataAtualizacao = new Date(ano, mes - 1, dia);
    return Math.floor((HOJE - dataAtualizacao) / (1000 * 60 * 60 * 24));
}

// Executando o código
main().catch((err) => {
    console.error("Erro na execução do script: ", err);
});