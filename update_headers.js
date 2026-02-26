const fs = require('fs');

try {
    const images = JSON.parse(fs.readFileSync('images_base64.json', 'utf8'));
    const imgLeft = images.left;
    const imgRight = images.right;

    // --- Update print.html ---
    let htmlPrint = fs.readFileSync('print.html', 'utf8');
    const newHeaderPrint = `        <div class="header-relatorio" style="display: flex; justify-content: space-between; align-items: center; border-bottom: 2px solid #000; margin-bottom: 20px; padding-bottom: 10px;">
            <img class="logo-left" src="data:image/png;base64,${imgLeft}" alt="Logo Ceará">
            <div style="text-align: center;">
                <div class="logo-placeholder">RELATÓRIO DE PLANTÃO</div>
                <div>Departamento de Polícia do Interior Sul - DPI SUL</div>
                <div id="infoGeral" class="mt-2"></div>
            </div>
            <img class="logo-right" src="data:image/png;base64,${imgRight}" alt="Brasão PCCE">
        </div>`;
    
    // Replace the old header-relatorio div (match with possible attributes)
    htmlPrint = htmlPrint.replace(/<div class="header-relatorio"[\s\S]*?<\/div>/, newHeaderPrint);
    fs.writeFileSync('print.html', htmlPrint);
    console.log('print.html updated');

    // --- Update extra.html ---
    let htmlExtra = fs.readFileSync('extra.html', 'utf8');
    const newHeaderExtra = `    <div class="header-relatorio-extra" style="display: flex; justify-content: space-between; align-items: center; border-bottom: 2px solid black; margin-bottom: 20px; padding-bottom: 10px;">
        <img class="logo-left" src="data:image/png;base64,${imgLeft}" alt="Logo Ceará">
        <div style="text-align: center;">
            <div style="font-weight: bold; font-size: 14px; text-transform: uppercase;">Relatório de Serviço Extraordinário</div>
            <div id="nomeUnidadeSuperior" style="font-size: 12px; font-weight: bold; text-transform: uppercase;"></div>
        </div>
        <img class="logo-right" src="data:image/png;base64,${imgRight}" alt="Brasão PCCE">
    </div>`;

    // Replace the header-relatorio-extra div (match with possible attributes)
    htmlExtra = htmlExtra.replace(/<div class="header-relatorio-extra"[\s\S]*?<\/div>/, newHeaderExtra);
    fs.writeFileSync('extra.html', htmlExtra);
    console.log('extra.html updated');

} catch (err) {
    console.error('Error updating files:', err);
    process.exit(1);
}
