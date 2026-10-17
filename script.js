let registros = [];

fetch("Residencias.xlsx")
    .then(response => response.arrayBuffer())
    .then(data => {

        const workbook = XLSX.read(data, {
            type: "array"
        });

        const sheetName = workbook.SheetNames[0];

        const worksheet = workbook.Sheets[sheetName];

        registros = XLSX.utils.sheet_to_json(worksheet);

        console.log("Planilha carregada");
        console.log(registros);

        alert("Registros carregados: " + registros.length);
    })
    .catch(error => {
        console.error(error);
        alert("Erro ao carregar planilha");
    });

function buscarRegistro(){

    const codigo =
        document.getElementById("codigo")
        .value
        .trim();

    const registro =
        registros.find(
            item => String(item.Codigo) === codigo
        );

    const resultado =
        document.getElementById("resultado");

    if(!registro){

        resultado.innerHTML =
            "<strong>Código não encontrado.</strong>";

        return;
    }

    const latitude =
        String(registro.Latitude)
        .replace(",", ".");

    const longitude =
        String(registro.Longitude)
        .replace(",", ".");

    const maps =
        `https://www.google.com/maps?q=${latitude},${longitude}`;

    resultado.innerHTML = `
        <p><strong>Código:</strong> ${registro.Codigo}</p>
        <p><strong>Latitude:</strong> ${latitude}</p>
        <p><strong>Longitude:</strong> ${longitude}</p>

        <a class="botao-mapa"
           href="${maps}"
           target="_blank">
           📍 Abrir Google Maps
        </a>
    `;
}
