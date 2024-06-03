let pFDV = document.getElementById('FDVDrop')
let dFDV = document.getElementById('FDVOptions')
let pMP = document.getElementById('MapaParqueDrop')
let dMP = document.getElementById('MapaParqueOptions')

let as = document.querySelectorAll('a')

as.forEach(e => {
    if(!e.href) {
        e.classList.add('active')
    } else if(e.href || e.href.length === 7) {
        e.classList.add('desactive')
    }
})

pFDV.addEventListener('click', () => {
    dFDV.classList.toggle('hide')
})

pMP.addEventListener('click', () => {
    dMP.classList.toggle('hide')
})