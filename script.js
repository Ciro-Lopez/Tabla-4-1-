fetch('Predicciones mundial 2026 (Respuestas).xlsx?v=' + Date.now())
    .then(res => res.arrayBuffer())
    .then(buffer => {
        const data = new Uint8Array(buffer);

        const workbook = XLSX.read(data, { type: 'array' });

        const hoja = workbook.Sheets[workbook.SheetNames[0]];

        const filas = XLSX.utils.sheet_to_json(hoja, { header: 1, defval: '' });

        const resultados = filas[20];

        const usuarios = [];
        for (let i = 1; i < filas.length; i++) {
            const fila = filas[i];

            if (i === 20) continue;

            const nombre = fila[2];

            if (!nombre || nombre.trim() === '') continue;

            let puntos = 0;

            for (let col = 3; col < resultados.length; col++) {
                if (resultados[col] !== '' && fila[col] === resultados[col]) {
                    puntos++;
                }
            }

            usuarios.push({ nombre, puntos });
        }

        usuarios.sort((a, b) => b.puntos - a.puntos);

        const tbody = document.getElementById('tabla-body');
        tbody.innerHTML = '';

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