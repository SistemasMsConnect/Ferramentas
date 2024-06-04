const arquivo = document.getElementById('fileInput')
const btn = document.getElementById('btn')
const fileName = document.getElementById('labelFileInput')
const pProcessando = document.getElementById('pProcessando')

document.getElementById('fileInput').addEventListener('change', function (event) {
    console.time('meu timer')
    const file = event.target.files[0];

    if (!file) {
        return;
    }

    const reader = new FileReader();

    reader.onload = function (e) {
        const content = e.target.result;
        // Processar o conteúdo do arquivo
        processCSV(content);
    };

    reader.readAsText(file, 'ISO-8859-1');

    pProcessando.setAttribute('style', 'display: inline-block')
    btn.setAttribute('style', 'display: none')
    fileName.textContent = file.name
});

function processCSV(content) {
    // Divida o conteúdo em linhas
    const lines = content.split('\n');

    let newCSVContent = '';

    // Pegar a primeira linha (cabeçalho)
    const firstLine = lines.shift();
    const splitFirstLine = firstLine.split(';')

    splitFirstLine.splice(26, 9)
    splitFirstLine.splice(24, 1)
    splitFirstLine.splice(21, 2)
    splitFirstLine.splice(19, 1)
    splitFirstLine.splice(7, 9)
    splitFirstLine.splice(4, 1)
    splitFirstLine.splice(1, 1)

    splitFirstLine[1] = 'REDE'

    const firstLineIndexTwo = splitFirstLine.slice(1, 2)
    const string = firstLineIndexTwo[0]

    splitFirstLine.splice(5, 0, string)
    splitFirstLine.splice(1, 1)
    const joinFirstLine = splitFirstLine.join(',') + '\n'


    // Remove a ultima linha em branco
    let lastLine = lines.length - 1
    if (lines[lastLine] == "") {
        lines.splice(lastLine, 1)
    }

    // Crie uma variável para armazenar o conteúdo do novo arquivo CSV
    let csvData = [];

    // Processar cada linha
    lines.forEach(function (line) {
        // Divida cada linha em colunas
        const columns = line.split(';');

        // 1. Ajustar linhas que foram empurradas erroniamente
        if (columns[35] === '' || columns[35] === null || columns[35] === undefined) {

        } else {
            columns.splice(10, 1)
        }

        // 2. Excluir colunas
        columns.splice(26, 9)
        columns.splice(24, 1)
        columns.splice(21, 2)
        columns.splice(19, 1)
        columns.splice(7, 9)
        columns.splice(4, 1)
        columns.splice(1, 1)

        // 3. Adicionar 0 nas células vazias
        if (columns[0].length < 5 || columns[0] === undefined) {
            columns[0] = '0'
        }

        if (columns[1].length < 5 || columns[1] === undefined) {
            columns[1] = '0'
        }

        if (columns[2].length < 5 || columns[2] === undefined) {
            columns[2] = '0'
        }

        if (columns[3].length <= 1) {
            if (columns[3] === '0' || columns[3] === undefined || columns[3] === null || columns[3] === '') {
                columns[3] = '0'
            }
        }

        if (columns[4].length <= 2 || columns[4] === undefined) {
            columns[4] = '0'
        }

        if (columns[5].length <= 2 || columns[5] === undefined) {
            columns[5] = '0'
        }

        if (columns[6].length <= 2 || columns[6] === undefined) {
            columns[6] = '0'
        }

        if (columns[7].length <= 2 || columns[7] === undefined) {
            columns[7] = '0'
        }

        if (columns[8].length < 5 || columns[8] === undefined) {
            columns[8] = 'zzz'
        }

        if (columns[10].length <= 2 || columns[10] === undefined) {
            columns[10] = '0'
        }

        csvData.push({
            col1: columns[0],
            col2: columns[2],
            col3: columns[3],
            col4: columns[4],
            col5: columns[1],
            col6: columns[5],
            col7: columns[6],
            col8: columns[7],
            col9: columns[8], // Coluna 9
            col10: columns[9],
            col11: columns[10],
        });
    });

    csvData.sort((a, b) => {
        return a.col9.localeCompare(b.col9);
    });

    csvData.forEach((e) => {
        if (e.col9 === 'zzz') {
            e.col9 = '0'
        }

        let dados = Object.values(e)
        dados.forEach((e) => {
            if (e.endsWith(' ')) {
                e = e.trim()
            }
        })

        if (e.col9 === 'BDL_IP_FIXO') {
            e.col2 += ' ' + e.col9
        }

        if (e.col8 === 'FTTH') {
            e.col11 = 'BANDA LARGA - FIBRA'
        } else if (e.col8 === 'VoIP FTTH') {
            e.col11 = 'TERMINAL - FIBRA'
        } else if (e.col8 === 'FTTC' || e.col8 === 'aDSL') {
            e.col11 = 'BANDA LARGA - METALICO'
        } else if (e.col8 === 'COBRE' || e.col8 === 'WLL') {
            e.col11 = 'TERMINAL - METALICO'
        } else if (e.col8 === 'IPTV') {
            e.col11 = 'TV'
        } else {
            e.col11 = 'OUTROS'
        }

        e.col3 = parseFloat(e.col3)
    })

    // Ordenar
    csvData.sort((b, a) => {
        return a.col10.localeCompare(b.col10);
    });

    csvData.sort((a, b) => {
        return a.col1.localeCompare(b.col1);
    });

    // Função para unificar os objetos e somar a chave 4 quando a chave 3 for igual
    const resultado = csvData.reduce((acc, obj) => {
        const chave = obj.col2;
        if (!acc[chave]) {
            acc[chave] = { ...obj }; // Cria uma cópia do objeto
        } else {
            acc[chave].col3 += parseFloat(obj.col3); // Soma a chave 4
        }
        return acc;
    }, {});

    // Funções de comparação para ordenar
    const compare = (b, a) => {
        if (a.col10 < b.col10) {
            return -1
        }

        if (a.col10 > b.col10) {
            return 1
        }
        return 0
    }

    const compare2 = (a, b) => {
        if (a.col1 < b.col1) {
            return -1
        }
        if (a.col1 > b.col1) {
            return 1
        }
        return 0
    }

    // Ordenar
    const certo1 = Object.values(resultado).sort(compare)
    const certo2 = certo1.sort(compare2)

    // Mantêm apenas duas casas decimais
    certo2.forEach((e) => {
        e.col3 = e.col3.toFixed(2)
    })

    // Removendo acentuações
    certo2.forEach((e) => {
        if (e.col4.includes('Móvel')) {
            e.col4 = e.col4.replace('Móvel', 'Movel')
        }

        if (e.col4.includes('Básico')) {
            e.col4 = e.col4.replace('Básico', 'Basico')
        }

        if (e.col4.includes('Escrit¢rio')) {
            e.col4 = e.col4.replace('Escrit¢rio', 'Escritorio')
        }

        if (e.col4.includes('EscritÂ¢rio')) {
            e.col4 = e.col4.replace('EscritÂ¢rio', 'Escritorio')
        }

        if (e.col4.includes('B sico')) {
            e.col4 = e.col4.replace('B sico', 'Basico')
        }
    })

    newCSVContent += joinFirstLine

    certo2.forEach(function (row) {
        newCSVContent += Object.values(row).join(',') + '\n';
    });

    // Criar um novo arquivo Blob
    const newCSVBlob = new Blob([newCSVContent], { type: 'text/csv' });

    // Criar um link de download para o novo arquivo
    const downloadLink = document.createElement('a');
    downloadLink.href = URL.createObjectURL(newCSVBlob);
    
    downloadLink.download = 'newFile.csv';

    // Adicionar o link ao corpo do documento
    document.body.appendChild(downloadLink);

    console.log('Terminou')
    console.timeEnd('meu timer')

    pProcessando.setAttribute('style', 'display: none')

    downloadLink.click();    

    btn.setAttribute('style', 'display: block; border: none; background-color: cadetblue; border-radius: 5px; color: aliceblue; cursor: pointer; padding: 5px 10px; margin-top: 15px')
    btn.addEventListener('click', () => downloadLink.click())
}