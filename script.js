fetch('16avos.xlsx?v=' + Date.now())
    .then(res => res.arrayBuffer())
    .then(buffer => {
        const data = new Uint8Array(buffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const hoja = workbook.Sheets[workbook.SheetNames[0]];
        const filas = XLSX.utils.sheet_to_json(hoja, { header: 1, defval: '' });

        const usuarios = [];

        for (let i = 1; i < filas.length; i++) {
            const fila = filas[i];

            const nombreRaw = fila[0];
            if (!nombreRaw || String(nombreRaw).trim() === '' || String(nombreRaw).trim() === 'Resultado') continue;

            // Puntos totales está en la penúltima columna (AY), no en la última (AZ)
            const puntosRaw = fila[fila.length - 2];
            const puntos = Number(puntosRaw);

            if (isNaN(puntos)) continue;

            usuarios.push({ nombre: String(nombreRaw).trim(), puntos });
        }

        usuarios.sort((a, b) => b.puntos - a.puntos);

        const tbody = document.getElementById('tabla-body');
        tbody.innerHTML = '';

        if (usuarios.length === 0) {
            tbody.innerHTML =
                '<tr><td colspan="3" style="text-align:center; color:#c0392b; padding:40px;">No se encontraron puntos válidos en el archivo</td></tr>';
            return;
        }

        usuarios.forEach((u, index) => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${index + 1}</td>
                <td>${u.nombre}</td>
                <td>${u.puntos}</td>
            `;
            tbody.appendChild(tr);
        });
    })
    .catch(() => {
        document.getElementById('tabla-body').innerHTML =
            '<tr><td colspan="3" style="text-align:center; color:#c0392b; padding:40px;">No se encontró el archivo Excel</td></tr>';
    });