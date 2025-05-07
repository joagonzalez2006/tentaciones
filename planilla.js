

let datosExcel = [];

document.getElementById('archivoExcel').addEventListener('change', function(e) {
  const archivo = e.target.files[0];
  const lector = new FileReader();

  lector.onload = function(e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });

    // Lee la primera hoja
    const hoja = workbook.Sheets[workbook.SheetNames[0]];
    datosExcel = XLSX.utils.sheet_to_json(hoja, { header: 1 });

    alert('Archivo cargado correctamente.');
  };

  lector.readAsArrayBuffer(archivo);
});

function buscar() {
  const entrada = document.getElementById('busqueda').value.trim();
  const resultadoDiv = document.getElementById('resultado');

  if (datosExcel.length === 0) {
    resultadoDiv.innerHTML = "<p>Primero debes cargar un archivo.</p>";
    return;
  }

  // Busca desde la fila 2, ya que la 1 tiene los encabezados
  const encontrados = datosExcel.slice(1).filter(fila => {
    if (!fila || fila.length === 0) return false;
    
    // Asegurarse de que los valores existan y convertirlos a string
    const codigo = fila[0] !== undefined ? String(fila[0]) : '';
    const nombre = fila[1] !== undefined ? String(fila[1]).toLowerCase() : '';
    
    // Buscar coincidencia exacta por código o parcial por nombre
    return codigo === entrada || nombre.includes(entrada.toLowerCase());
  });

  if (encontrados.length > 0) {
    // Crear un cuadro de información para mostrar los resultados
    let resultadoHTML = '<div class="resultado-cuadro">';
    
    encontrados.forEach(fila => {
      const codigo = fila[0] || 'N/A';
      const nombre = fila[1] || 'N/A';
      // La columna D es el índice 3 (0-indexed)
      const precio = fila[3] !== undefined ? fila[3] : 'N/A';
      
      resultadoHTML += `
        <div class="producto-info">
          <p><strong>Código:</strong> ${codigo}</p>
          <p><strong>Nombre:</strong> ${nombre}</p>
          <p><strong>Precio:</strong> $ ${precio}</p>
        </div>
      `;
    });
    
    resultadoHTML += '</div>';
    resultadoDiv.innerHTML = resultadoHTML;
  } else {
    resultadoDiv.innerHTML = "<p>No se encontraron resultados.</p>";
  }
}