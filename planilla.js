let datosExcel = [];
let hojas = [];

document.getElementById('archivoExcel').addEventListener('change', function(e) {
  const archivo = e.target.files[0];
  const lector = new FileReader();

  lector.onload = function(e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    
    // Guardar nombres de todas las hojas
    hojas = workbook.SheetNames;
    
    // Limpiar y crear selector de hojas
    const selectorHojas = document.getElementById('selectorHojas');
    selectorHojas.innerHTML = '';
    selectorHojas.style.display = 'block';
    
    // Agregar cada hoja como opción
    hojas.forEach(nombreHoja => {
      const opcion = document.createElement('option');
      opcion.value = nombreHoja;
      opcion.textContent = nombreHoja;
      selectorHojas.appendChild(opcion);
    });
    
    // Cargar la primera hoja por defecto
    cargarHoja(hojas[0], workbook);
    
    document.getElementById('infoHojas').style.display = 'block';
    alert('Archivo cargado correctamente. Se encontraron ' + hojas.length + ' hojas.');
  };

  lector.readAsArrayBuffer(archivo);
});

// Función para cargar una hoja específica
function cargarHoja(nombreHoja, workbook) {
  const hoja = workbook.Sheets[nombreHoja];
  datosExcel = XLSX.utils.sheet_to_json(hoja, { header: 1 });
  document.getElementById('hojaActual').textContent = nombreHoja;
  document.getElementById('resultado').innerHTML = '';
}

// Función para cambiar de hoja
document.getElementById('selectorHojas').addEventListener('change', function() {
  const nombreHoja = this.value;
  
  // Necesitamos volver a leer el archivo
  const archivo = document.getElementById('archivoExcel').files[0];
  if (!archivo) return;
  
  const lector = new FileReader();
  lector.onload = function(e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    cargarHoja(nombreHoja, workbook);
  };
  
  lector.readAsArrayBuffer(archivo);
});

function buscar() {
  const entrada = document.getElementById('busqueda').value.trim();
  const tipoBusqueda = document.getElementById('tipoBusqueda').value;
  const resultadoDiv = document.getElementById('resultado');

  if (datosExcel.length === 0) {
    resultadoDiv.innerHTML = "<p>Primero debes cargar un archivo.</p>";
    return;
  }

  // Definimos las columnas con índices fijos
  const indiceCodigoArticulo = 0; // Columna A
  const indiceNombre = 1;         // Columna B
  const indiceCodigoOrigen = 2;   // Columna C
  const indicePrecio = 3;         // Columna D

  // Busca desde la fila 2, ya que la 1 tiene los encabezados
  const encontrados = datosExcel.slice(1).filter(fila => {
    if (!fila || fila.length === 0) return false;
    
    // Asegurarse de que los valores existan y convertirlos a string
    const codigoArticulo = fila[indiceCodigoArticulo] !== undefined ? String(fila[indiceCodigoArticulo]) : '';
    const nombre = fila[indiceNombre] !== undefined ? String(fila[indiceNombre]).toLowerCase() : '';
    const codigoOrigen = fila[indiceCodigoOrigen] !== undefined ? String(fila[indiceCodigoOrigen]) : '';
    
    // Buscar según el tipo de búsqueda seleccionado
    switch (tipoBusqueda) {
      case 'codigo':
        return codigoArticulo === entrada;
      case 'origen':
        return codigoOrigen.toLowerCase() === entrada.toLowerCase();
      case 'nombre':
        return nombre.includes(entrada.toLowerCase());
      default: // Búsqueda en todos los campos
        return codigoArticulo === entrada || 
               nombre.includes(entrada.toLowerCase()) || 
               codigoOrigen.toLowerCase() === entrada.toLowerCase();
    }
  });

  if (encontrados.length > 0) {
    // Crear un cuadro de información para mostrar los resultados
    let resultadoHTML = '<div class="resultado-cuadro">';
    
    encontrados.forEach(fila => {
      const codigoArticulo = fila[indiceCodigoArticulo] || 'N/A';
      const nombre = fila[indiceNombre] || 'N/A';
      const codigoOrigen = fila[indiceCodigoOrigen] || 'N/A';
      const precio = fila[indicePrecio] !== undefined ? fila[indicePrecio] : 'N/A';
      
      resultadoHTML += `
        <div class="producto-info">
          <p><strong>Código artículo:</strong> ${codigoArticulo}</p>
          <p><strong>Nombre:</strong> ${nombre}</p>
          <p><strong>Código de origen:</strong> ${codigoOrigen}</p>
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
