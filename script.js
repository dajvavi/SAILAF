document.addEventListener('DOMContentLoaded', function() {
    // Buscar el botón "Capturar OT" según su texto
    const botones = document.querySelectorAll('.header-buttons button');
    let capturarOTButton = null;
    botones.forEach(button => {
      if (button.textContent.trim() === "Capturar OT") {
        capturarOTButton = button;
      }
    });
  
    if (capturarOTButton) {
      capturarOTButton.addEventListener('click', function() {
        // Creamos un input para seleccionar archivos Excel
        let fileInput = document.createElement('input');
        fileInput.type = 'file';
        fileInput.multiple = true;
        fileInput.accept = '.xls,.xlsx';
        fileInput.addEventListener('change', handleFiles, false);
        fileInput.click();
      });
    }
  
    // Obtenemos el <tbody> del primer table (donde se pegarán los números de muestra)
    const mainTableTbody = document.querySelector('.excel-container table tbody');
  
    function handleFiles(event) {
      const files = event.target.files;
      if (files.length === 0) {
        alert("No se ha seleccionado ningún archivo.");
        return;
      }
      // Si el <tbody> ya tiene datos, se impide continuar
      if (mainTableTbody.children.length > 0) {
        alert("La hoja ya tiene datos, no se puede continuar");
        return;
      }
  
      let allSamples = [];
      let fileIndex = 0;
  
      function processNext() {
        if (fileIndex < files.length) {
          processFile(files[fileIndex], function(samples) {
            // Acumulamos los números de muestra del archivo actual
            allSamples = allSamples.concat(samples);
            fileIndex++;
            processNext();
          });
        } else {
          // Una vez procesados todos los archivos, se actualiza la tabla
          updateTable(allSamples);
          alert("Se ha terminado la captura");
        }
      }
      processNext();
    }
  
    function processFile(file, callback) {
      let reader = new FileReader();
      reader.onload = function(e) {
        let data = new Uint8Array(e.target.result);
        let workbook = XLSX.read(data, { type: 'array' });
        let sheetName = "Sheet1";
        if (!workbook.Sheets[sheetName]) {
          alert(`El archivo ${file.name} no contiene la hoja "Sheet1".`);
          callback([]);
          return;
        }
        let worksheet = workbook.Sheets[sheetName];
        // Convertir la hoja en array de arrays (la primera fila es el encabezado)
        let jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        if (jsonData.length < 2) {
          alert(`El archivo ${file.name} está vacío o no tiene datos suficientes.`);
          callback([]);
          return;
        }
        let header = jsonData[0];
        let dataRows = jsonData.slice(1);
  
        // Buscar el índice de la columna "Fósforo soluble en bicarbonato sódico (OLSEN) (P)"
        let ncol = header.indexOf("Fósforo soluble en bicarbonato sódico (OLSEN) (P)");
        if (ncol === -1) {
          alert("No hay columna con muestras pendientes de análisis para esta determinación en " + file.name);
          callback([]);
          return;
        }
  
        let samples = [];
        // Recorremos las filas y, si el valor en la columna indicada es 1, capturamos el primer dato (número de muestra)
        for (let i = 0; i < dataRows.length && samples.length < 49; i++) {
          let row = dataRows[i];
          if (row[ncol] === 1 || row[ncol] === "1") {
            samples.push(row[0]); // Se asume que el primer valor es el número de muestra
          }
        }
  
        if (samples.length === 0) {
          alert("No hay muestras pendientes en " + file.name);
        }
        if (samples.length >= 50) {
          alert("Has intentado capturar demasiadas muestras en " + file.name + ", el listado se limitará a 49");
        }
        callback(samples);
      };
      reader.readAsArrayBuffer(file);
    }
  
    function updateTable(sampleNumbers) {
      // Limpiar el contenido previo del <tbody>
      mainTableTbody.innerHTML = "";
      // Insertar una fila por cada número de muestra en la columna "MUESTRA"
      sampleNumbers.forEach(function(sample) {
        let row = document.createElement('tr');
        // La primera celda contiene el número de muestra; el resto se dejan vacíos
        row.innerHTML = `<td>${sample}</td><td></td><td></td><td></td><td></td><td></td>`;
        mainTableTbody.appendChild(row);
      });
  
      // Calcular y agregar las filas de repeticiones al final
      let n = sampleNumbers.length;
      if (n > 0) {
        let j = Math.ceil(n / 10);
        // Primera repetición: se toma el primer número de muestra (simulando ws.Range("A3").Value)
        let repRow = document.createElement('tr');
        repRow.innerHTML = `<td>R-${sampleNumbers[0]}</td><td></td><td></td><td></td><td></td><td></td>`;
        mainTableTbody.appendChild(repRow);
        // Si hay más de 20 muestras, se agregan repeticiones intermedias
        if (j > 2) {
          for (let w = 1; w <= (j - 2); w++) {
            let index = Math.ceil((n / j) * w) - 1;
            if (index < sampleNumbers.length) {
              let repRow = document.createElement('tr');
              repRow.innerHTML = `<td>R-${sampleNumbers[index]}</td><td></td><td></td><td></td><td></td><td></td>`;
              mainTableTbody.appendChild(repRow);
            }
          }
        }
        // Última repetición: se toma el último número de muestra
        if (j > 1) {
          let repRow2 = document.createElement('tr');
          repRow2.innerHTML = `<td>R-${sampleNumbers[n - 1]}</td><td></td><td></td><td></td><td></td><td></td>`;
          mainTableTbody.appendChild(repRow2);
        }
      }
    }
  });
 
  document.addEventListener('DOMContentLoaded', function() {
    // Buscamos el botón para borrar datos por su id
    const clearDataButton = document.getElementById('clear-data');
  
    clearDataButton.addEventListener('click', function() {
      // Pedir confirmación al usuario
      const confirmar = confirm("¿Estás seguro de que deseas borrar los datos de muestra?");
      if (confirmar) {
        // Selecciona el <tbody> del primer table donde se encuentran los datos
        const mainTableTbody = document.querySelector('.excel-container table tbody');
        // Limpia el contenido
        mainTableTbody.innerHTML = "";
        alert("Datos de muestra borrados.");
      }
    });
  });
  