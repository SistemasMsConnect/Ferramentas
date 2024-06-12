const fileInput = document.getElementById('fileInput')
const exportBtn = document.getElementById('exportBtn')
const pProcessando = document.getElementById('p')

fileInput.addEventListener('change', (arquivo) => {
    var file = arquivo.target.files[0]

    pProcessando.classList.remove('hide')

    var reader = new FileReader();
    reader.onload = function (event) {
        var data = new Uint8Array(event.target.result);
        var workbook = XLSX.read(data, { type: 'array' });

        var sheetName01 = workbook.SheetNames[0];
        var sheetName02 = workbook.SheetNames[1];
        var sheetName03 = workbook.SheetNames[2];
        var sheetName04 = workbook.SheetNames[3];
        var sheetName05 = workbook.SheetNames[4];
        var sheetName06 = workbook.SheetNames[5];

        var sheet01 = workbook.Sheets[sheetName01];
        var sheet02 = workbook.Sheets[sheetName02];
        var sheet03 = workbook.Sheets[sheetName03];
        var sheet04 = workbook.Sheets[sheetName04];
        var sheet05 = workbook.Sheets[sheetName05];
        var sheet06 = workbook.Sheets[sheetName06];

        let jsonData01 = XLSX.utils.sheet_to_json(sheet01, { defval: 0 });
        let jsonData02 = XLSX.utils.sheet_to_json(sheet02, { defval: 0 });
        let jsonData03 = XLSX.utils.sheet_to_json(sheet03, { defval: 0 });
        let jsonData04 = XLSX.utils.sheet_to_json(sheet04, { defval: 0 });
        let jsonData05 = XLSX.utils.sheet_to_json(sheet05, { defval: 0 });
        let jsonData06 = XLSX.utils.sheet_to_json(sheet06, { defval: 0 });

        processarAbaUm(jsonData01)
        processarAbaDois(jsonData02)
        processarAbaTres(jsonData03)
        processarAbaQuatro(jsonData04)
        processarAbaCinco(jsonData05)
        processarAbaSeis(jsonData06)

        createFile(avancados1Um, avancados1Dois, avancados1Tres, avancados1Quatro, avancados1Cinco, avancados1Seis, "avancada1")
        createFile(avancados2Um, avancados2Dois, avancados2Tres, avancados2Quatro, avancados2Cinco, avancados2Seis, "avancada2")
        createFile(cpfUm, cpfDois, cpfTres, cpfQuatro, cpfCinco, cpfSeis, "cpf")
        createFile(basica1Um, basica1Dois, basica1Tres, basica1Quatro, basica1Cinco, basica1Seis, "basica1")
        createFile(basica2Um, basica2Dois, basica2Tres, basica2Quatro, basica2Cinco, basica2Seis, "basica2")
        createFile(basica3Um, basica3Dois, basica3Tres, basica3Quatro, basica3Cinco, basica3Seis, "basica3")
        createFile(basica4Um, basica4Dois, basica4Tres, basica4Quatro, basica4Cinco, basica4Seis, "basica4")
        createFile(basica5Um, basica5Dois, basica5Tres, basica5Quatro, basica5Cinco, basica5Seis, "basica5")
        createFile(basica6Um, basica6Dois, basica6Tres, basica6Quatro, basica6Cinco, basica6Seis, "basica6")
        createFile(movel1Um, movel1Dois, movel1Tres, movel1Quatro, movel1Cinco, movel1Seis, "movel1")
        createFile(movel2Um, movel2Dois, movel2Tres, movel2Quatro, movel2Cinco, movel2Seis, "movel2")
        createFile(movel3Um, movel3Dois, movel3Tres, movel3Quatro, movel3Cinco, movel3Seis, "movel3")
        createFile(movel4Um, movel4Dois, movel4Tres, movel4Quatro, movel4Cinco, movel4Seis, "movel4")
        createFile(movel5Um, movel5Dois, movel5Tres, movel5Quatro, movel5Cinco, movel5Seis, "movel5")
        createFile(movel6Um, movel6Dois, movel6Tres, movel6Quatro, movel6Cinco, movel6Seis, "movel6")
        createFile(mb1Um, mb1Dois, mb1Tres, mb1Quatro, mb1Cinco, mb1Seis, "mb1")
        createFile(mb2Um, mb2Dois, mb2Tres, mb2Quatro, mb2Cinco, mb2Seis, "mb2")
        createFile(mb3Um, mb3Dois, mb3Tres, mb3Quatro, mb3Cinco, mb3Seis, "mb3")
        createFile(mb4Um, mb4Dois, mb4Tres, mb4Quatro, mb4Cinco, mb4Seis, "mb4")
        createFile(mb5Um, mb5Dois, mb5Tres, mb5Quatro, mb5Cinco, mb5Seis, "mb5")
        createFile(mb6Um, mb6Dois, mb6Tres, mb6Quatro, mb6Cinco, mb6Seis, "mb6")

        pProcessando.classList.add('hide')

        console.log(jsonData01)
        console.log(jsonData02)
        console.log(jsonData03)
        console.log(jsonData04)
        console.log(jsonData05)
        console.log(jsonData06)
    };

    reader.readAsArrayBuffer(file);
})


let avancados1Um = []
let avancados2Um = []
let cpfUm = []
let basica1Um = []
let basica2Um = []
let basica3Um = []
let basica4Um = []
let basica5Um = []
let basica6Um = []
let movel1Um = []
let movel2Um = []
let movel3Um = []
let movel4Um = []
let movel5Um = []
let movel6Um = []
let mb1Um = []
let mb2Um = []
let mb3Um = []
let mb4Um = []
let mb5Um = []
let mb6Um = []

let avancados1Dois = []
let avancados2Dois = []
let cpfDois = []
let basica1Dois = []
let basica2Dois = []
let basica3Dois = []
let basica4Dois = []
let basica5Dois = []
let basica6Dois = []
let movel1Dois = []
let movel2Dois = []
let movel3Dois = []
let movel4Dois = []
let movel5Dois = []
let movel6Dois = []
let mb1Dois = []
let mb2Dois = []
let mb3Dois = []
let mb4Dois = []
let mb5Dois = []
let mb6Dois = []

let avancados1Tres = []
let avancados2Tres = []
let cpfTres = []
let basica1Tres = []
let basica2Tres = []
let basica3Tres = []
let basica4Tres = []
let basica5Tres = []
let basica6Tres = []
let movel1Tres = []
let movel2Tres = []
let movel3Tres = []
let movel4Tres = []
let movel5Tres = []
let movel6Tres = []
let mb1Tres = []
let mb2Tres = []
let mb3Tres = []
let mb4Tres = []
let mb5Tres = []
let mb6Tres = []

let avancados1Quatro = []
let avancados2Quatro = []
let cpfQuatro = []
let basica1Quatro = []
let basica2Quatro = []
let basica3Quatro = []
let basica4Quatro = []
let basica5Quatro = []
let basica6Quatro = []
let movel1Quatro = []
let movel2Quatro = []
let movel3Quatro = []
let movel4Quatro = []
let movel5Quatro = []
let movel6Quatro = []
let mb1Quatro = []
let mb2Quatro = []
let mb3Quatro = []
let mb4Quatro = []
let mb5Quatro = []
let mb6Quatro = []

let avancados1Cinco = []
let avancados2Cinco = []
let cpfCinco = []
let basica1Cinco = []
let basica2Cinco = []
let basica3Cinco = []
let basica4Cinco = []
let basica5Cinco = []
let basica6Cinco = []
let movel1Cinco = []
let movel2Cinco = []
let movel3Cinco = []
let movel4Cinco = []
let movel5Cinco = []
let movel6Cinco = []
let mb1Cinco = []
let mb2Cinco = []
let mb3Cinco = []
let mb4Cinco = []
let mb5Cinco = []
let mb6Cinco = []

let avancados1Seis = []
let avancados2Seis = []
let cpfSeis = []
let basica1Seis = []
let basica2Seis = []
let basica3Seis = []
let basica4Seis = []
let basica5Seis = []
let basica6Seis = []
let movel1Seis = []
let movel2Seis = []
let movel3Seis = []
let movel4Seis = []
let movel5Seis = []
let movel6Seis = []
let mb1Seis = []
let mb2Seis = []
let mb3Seis = []
let mb4Seis = []
let mb5Seis = []
let mb6Seis = []


function processarAbaUm(content) {
    content.forEach(e => {
        if (e.N === 'CPF') {
            cpfUm.push(e)
        } else if (e.N === 'Avançados&TI M-B - 1') {
            avancados1Um.push(e)
        } else if (e.N === 'Avançados&TI M-B - 2') {
            avancados2Um.push(e)
        } else if (e.N === 'Basica-1') {
            basica1Um.push(e)
        } else if (e.N === 'Basica-2') {
            basica2Um.push(e)
        } else if (e.N === 'Basica-3') {
            basica3Um.push(e)
        } else if (e.N === 'Basica-4') {
            basica4Um.push(e)
        } else if (e.N === 'Basica-5') {
            basica5Um.push(e)
        } else if (e.N === 'Basica-6') {
            basica6Um.push(e)
        } else if (e.N === 'Movel-1') {
            movel1Um.push(e)
        } else if (e.N === 'Movel-2') {
            movel2Um.push(e)
        } else if (e.N === 'Movel-3') {
            movel3Um.push(e)
        } else if (e.N === 'Movel-4') {
            movel4Um.push(e)
        } else if (e.N === 'Movel-5') {
            movel5Um.push(e)
        } else if (e.N === 'Movel-6') {
            movel6Um.push(e)
        } else if (e.N === 'Movel&Basica-1') {
            mb1Um.push(e)
        } else if (e.N === 'Movel&Basica-2') {
            mb2Um.push(e)
        } else if (e.N === 'Movel&Basica-3') {
            mb3Um.push(e)
        } else if (e.N === 'Movel&Basica-4') {
            mb4Um.push(e)
        } else if (e.N === 'Movel&Basica-5') {
            mb5Um.push(e)
        } else if (e.N === 'Movel&Basica-6') {
            mb6Um.push(e)
        }

        // Deletar propriedades
        delete e.N
    })
}

function processarAbaDois(content) {
    content.forEach(e => {
        if (e.N === 'CPF') {
            cpfDois.push(e)
        } else if (e.N === 'Avançados&TI M-B - 1') {
            avancados1Dois.push(e)
        } else if (e.N === 'Avançados&TI M-B - 2') {
            avancados2Dois.push(e)
        } else if (e.N === 'Basica-1') {
            basica1Dois.push(e)
        } else if (e.N === 'Basica-2') {
            basica2Dois.push(e)
        } else if (e.N === 'Basica-3') {
            basica3Dois.push(e)
        } else if (e.N === 'Basica-4') {
            basica4Dois.push(e)
        } else if (e.N === 'Basica-5') {
            basica5Dois.push(e)
        } else if (e.N === 'Basica-6') {
            basica6Dois.push(e)
        } else if (e.N === 'Movel-1') {
            movel1Dois.push(e)
        } else if (e.N === 'Movel-2') {
            movel2Dois.push(e)
        } else if (e.N === 'Movel-3') {
            movel3Dois.push(e)
        } else if (e.N === 'Movel-4') {
            movel4Dois.push(e)
        } else if (e.N === 'Movel-5') {
            movel5Dois.push(e)
        } else if (e.N === 'Movel-6') {
            movel6Dois.push(e)
        } else if (e.N === 'Movel&Basica-1') {
            mb1Dois.push(e)
        } else if (e.N === 'Movel&Basica-2') {
            mb2Dois.push(e)
        } else if (e.N === 'Movel&Basica-3') {
            mb3Dois.push(e)
        } else if (e.N === 'Movel&Basica-4') {
            mb4Dois.push(e)
        } else if (e.N === 'Movel&Basica-5') {
            mb5Dois.push(e)
        } else if (e.N === 'Movel&Basica-6') {
            mb6Dois.push(e)
        }

        // Deletar propriedades
        delete e.N
    })
}

function processarAbaTres(content) {
    content.forEach(e => {
        if (e.N === 'CPF') {
            cpfTres.push(e)
        } else if (e.N === 'Avançados&TI M-B - 1') {
            avancados1Tres.push(e)
        } else if (e.N === 'Avançados&TI M-B - 2') {
            avancados2Tres.push(e)
        } else if (e.N === 'Basica-1') {
            basica1Tres.push(e)
        } else if (e.N === 'Basica-2') {
            basica2Tres.push(e)
        } else if (e.N === 'Basica-3') {
            basica3Tres.push(e)
        } else if (e.N === 'Basica-4') {
            basica4Tres.push(e)
        } else if (e.N === 'Basica-5') {
            basica5Tres.push(e)
        } else if (e.N === 'Basica-6') {
            basica6Tres.push(e)
        } else if (e.N === 'Movel-1') {
            movel1Tres.push(e)
        } else if (e.N === 'Movel-2') {
            movel2Tres.push(e)
        } else if (e.N === 'Movel-3') {
            movel3Tres.push(e)
        } else if (e.N === 'Movel-4') {
            movel4Tres.push(e)
        } else if (e.N === 'Movel-5') {
            movel5Tres.push(e)
        } else if (e.N === 'Movel-6') {
            movel6Tres.push(e)
        } else if (e.N === 'Movel&Basica-1') {
            mb1Tres.push(e)
        } else if (e.N === 'Movel&Basica-2') {
            mb2Tres.push(e)
        } else if (e.N === 'Movel&Basica-3') {
            mb3Tres.push(e)
        } else if (e.N === 'Movel&Basica-4') {
            mb4Tres.push(e)
        } else if (e.N === 'Movel&Basica-5') {
            mb5Tres.push(e)
        } else if (e.N === 'Movel&Basica-6') {
            mb6Tres.push(e)
        }

        // Deletar propriedades
        delete e.N
    })
}

function processarAbaQuatro(content) {
    content.forEach(e => {
        if (e.N === 'CPF') {
            cpfQuatro.push(e)
        } else if (e.N === 'Avançados&TI M-B - 1') {
            avancados1Quatro.push(e)
        } else if (e.N === 'Avançados&TI M-B - 2') {
            avancados2Quatro.push(e)
        } else if (e.N === 'Basica-1') {
            basica1Quatro.push(e)
        } else if (e.N === 'Basica-2') {
            basica2Quatro.push(e)
        } else if (e.N === 'Basica-3') {
            basica3Quatro.push(e)
        } else if (e.N === 'Basica-4') {
            basica4Quatro.push(e)
        } else if (e.N === 'Basica-5') {
            basica5Quatro.push(e)
        } else if (e.N === 'Basica-6') {
            basica6Quatro.push(e)
        } else if (e.N === 'Movel-1') {
            movel1Quatro.push(e)
        } else if (e.N === 'Movel-2') {
            movel2Quatro.push(e)
        } else if (e.N === 'Movel-3') {
            movel3Quatro.push(e)
        } else if (e.N === 'Movel-4') {
            movel4Quatro.push(e)
        } else if (e.N === 'Movel-5') {
            movel5Quatro.push(e)
        } else if (e.N === 'Movel-6') {
            movel6Quatro.push(e)
        } else if (e.N === 'Movel&Basica-1') {
            mb1Quatro.push(e)
        } else if (e.N === 'Movel&Basica-2') {
            mb2Quatro.push(e)
        } else if (e.N === 'Movel&Basica-3') {
            mb3Quatro.push(e)
        } else if (e.N === 'Movel&Basica-4') {
            mb4Quatro.push(e)
        } else if (e.N === 'Movel&Basica-5') {
            mb5Quatro.push(e)
        } else if (e.N === 'Movel&Basica-6') {
            mb6Quatro.push(e)
        }

        // Deletar propriedades
        delete e.N
    })
}

function processarAbaCinco(content) {
    content.forEach(e => {
        if (e.N === 'CPF') {
            cpfCinco.push(e)
        } else if (e.N === 'Avançados&TI M-B - 1') {
            avancados1Cinco.push(e)
        } else if (e.N === 'Avançados&TI M-B - 2') {
            avancados2Cinco.push(e)
        } else if (e.N === 'Basica-1') {
            basica1Cinco.push(e)
        } else if (e.N === 'Basica-2') {
            basica2Cinco.push(e)
        } else if (e.N === 'Basica-3') {
            basica3Cinco.push(e)
        } else if (e.N === 'Basica-4') {
            basica4Cinco.push(e)
        } else if (e.N === 'Basica-5') {
            basica5Cinco.push(e)
        } else if (e.N === 'Basica-6') {
            basica6Cinco.push(e)
        } else if (e.N === 'Movel-1') {
            movel1Cinco.push(e)
        } else if (e.N === 'Movel-2') {
            movel2Cinco.push(e)
        } else if (e.N === 'Movel-3') {
            movel3Cinco.push(e)
        } else if (e.N === 'Movel-4') {
            movel4Cinco.push(e)
        } else if (e.N === 'Movel-5') {
            movel5Cinco.push(e)
        } else if (e.N === 'Movel-6') {
            movel6Cinco.push(e)
        } else if (e.N === 'Movel&Basica-1') {
            mb1Cinco.push(e)
        } else if (e.N === 'Movel&Basica-2') {
            mb2Cinco.push(e)
        } else if (e.N === 'Movel&Basica-3') {
            mb3Cinco.push(e)
        } else if (e.N === 'Movel&Basica-4') {
            mb4Cinco.push(e)
        } else if (e.N === 'Movel&Basica-5') {
            mb5Cinco.push(e)
        } else if (e.N === 'Movel&Basica-6') {
            mb6Cinco.push(e)
        }

        // Deletar propriedades
        delete e.N
    })
}

function processarAbaSeis(content) {
    content.forEach(e => {
        if (e.N === 'CPF') {
            cpfSeis.push(e)
        } else if (e.N === 'Avançados&TI M-B - 1') {
            avancados1Seis.push(e)
        } else if (e.N === 'Avançados&TI M-B - 2') {
            avancados2Seis.push(e)
        } else if (e.N === 'Basica-1') {
            basica1Seis.push(e)
        } else if (e.N === 'Basica-2') {
            basica2Seis.push(e)
        } else if (e.N === 'Basica-3') {
            basica3Seis.push(e)
        } else if (e.N === 'Basica-4') {
            basica4Seis.push(e)
        } else if (e.N === 'Basica-5') {
            basica5Seis.push(e)
        } else if (e.N === 'Basica-6') {
            basica6Seis.push(e)
        } else if (e.N === 'Movel-1') {
            movel1Seis.push(e)
        } else if (e.N === 'Movel-2') {
            movel2Seis.push(e)
        } else if (e.N === 'Movel-3') {
            movel3Seis.push(e)
        } else if (e.N === 'Movel-4') {
            movel4Seis.push(e)
        } else if (e.N === 'Movel-5') {
            movel5Seis.push(e)
        } else if (e.N === 'Movel-6') {
            movel6Seis.push(e)
        } else if (e.N === 'Movel&Basica-1') {
            mb1Seis.push(e)
        } else if (e.N === 'Movel&Basica-2') {
            mb2Seis.push(e)
        } else if (e.N === 'Movel&Basica-3') {
            mb3Seis.push(e)
        } else if (e.N === 'Movel&Basica-4') {
            mb4Seis.push(e)
        } else if (e.N === 'Movel&Basica-5') {
            mb5Seis.push(e)
        } else if (e.N === 'Movel&Basica-6') {
            mb6Seis.push(e)
        }

        // Deletar propriedades
        delete e.N
    })
}

function createFile(data1, data2, data3, data4, data5, data6, name) {
    const wsNome1 = XLSX.utils.json_to_sheet(data1)
    const wsNome2 = XLSX.utils.json_to_sheet(data2)
    const wsNome3 = XLSX.utils.json_to_sheet(data3)
    const wsNome4 = XLSX.utils.json_to_sheet(data4)
    const wsNome5 = XLSX.utils.json_to_sheet(data5)
    const wsNome6 = XLSX.utils.json_to_sheet(data6)

    const workbook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(workbook, wsNome1, "Mapa Parque");
    XLSX.utils.book_append_sheet(workbook, wsNome2, "Parque móvel");
    XLSX.utils.book_append_sheet(workbook, wsNome3, "Parque Básica");
    XLSX.utils.book_append_sheet(workbook, wsNome4, "Parque de avançada");
    XLSX.utils.book_append_sheet(workbook, wsNome5, "Parque TI");
    XLSX.utils.book_append_sheet(workbook, wsNome6, "Recomendações");

    const wbout = XLSX.writeFile(workbook, `${name}.xlsx`, { compression: true });
    // Crie um blob com o conteúdo do arquivo XLSX
    const blob = new Blob([wbout], { type: 'application/octet-stream' });

    // Crie um URL para o blob
    const url = URL.createObjectURL(blob);

    // Crie um link de download e atribua o URL do blob como seu href
    const a = document.createElement('a');
    a.href = url;
    a.download = 'dados.xlsx';
    document.body.appendChild(a);

    // Simule um clique no link para iniciar o download
    a.click();
}