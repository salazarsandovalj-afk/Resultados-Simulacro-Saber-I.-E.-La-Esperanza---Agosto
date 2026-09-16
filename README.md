# Resultados-Simulacro-Saber-I.-E.-La-Esperanza---Agosto
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Institución Educativa La Esperanza — Dashboard Saber</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap" rel="stylesheet">
    <style>
        :root {
            --primary: #1e3a8a;
            --secondary: #2563eb;
            --accent: #10b981;
            --warning: #f59e0b;
            --danger: #ef4444;
            --bg-gray: #f8fafc;
            --card-bg: #ffffff;
            --text-main: #1e293b;
            --text-muted: #64748b;
        }

        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
            font-family: 'Inter', sans-serif;
        }

        body {
            background-color: var(--bg-gray);
            color: var(--text-main);
            padding: 24px;
        }

        header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 24px;
            background: linear-gradient(135deg, #1e3a8a, #2563eb);
            color: white;
            padding: 22px 30px;
            border-radius: 16px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.08);
        }

        header h1 { font-size: 1.6rem; font-weight: 700; }
        header p { font-size: 0.95rem; opacity: 0.9; margin-top: 4px; }

        .controls {
            display: flex;
            gap: 12px;
            margin-bottom: 24px;
        }

        select, button {
            padding: 10px 16px;
            border-radius: 8px;
            border: 1px solid #cbd5e1;
            background: white;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s;
        }

        button.btn-primary {
            background: var(--accent);
            color: white;
            border: none;
        }

        button.btn-primary:hover {
            background: #059669;
        }

        .metrics-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
            gap: 16px;
            margin-bottom: 24px;
        }

        .card {
            background: var(--card-bg);
            border-radius: 12px;
            padding: 20px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.04);
            border: 1px solid #e2e8f0;
        }

        .metric-title { font-size: 0.85rem; color: var(--text-muted); font-weight: 600; text-transform: uppercase; }
        .metric-value { font-size: 1.8rem; font-weight: 700; color: var(--primary); margin-top: 6px; }
        .metric-sub { font-size: 0.8rem; color: var(--text-muted); margin-top: 4px; }

        .dafo-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
            gap: 16px;
            margin-bottom: 24px;
        }

        .dafo-card {
            padding: 16px;
            border-radius: 10px;
            border-left: 5px solid;
            background: white;
            box-shadow: 0 2px 6px rgba(0,0,0,0.03);
        }

        .dafo-f { border-color: var(--accent); background: #f0fdf4; }
        .dafo-d { border-color: var(--danger); background: #fef2f2; }
        .dafo-o { border-color: var(--secondary); background: #eff6ff; }
        .dafo-a { border-color: var(--warning); background: #fffbeb; }

        .dafo-card h3 { font-size: 1rem; margin-bottom: 8px; }
        .dafo-card ul { font-size: 0.85rem; padding-left: 18px; line-height: 1.5; }

        .charts-grid {
            display: grid;
            grid-template-columns: 2fr 1fr;
            gap: 24px;
            margin-bottom: 24px;
        }

        @media (max-width: 900px) {
            .charts-grid { grid-template-columns: 1fr; }
        }

        .chart-card {
            background: var(--card-bg);
            border-radius: 12px;
            padding: 20px;
            border: 1px solid #e2e8f0;
        }

        .chart-card h2 { font-size: 1.1rem; margin-bottom: 16px; color: var(--primary); }

        table {
            width: 100%;
            border-collapse: collapse;
            font-size: 0.85rem;
        }

        th, td {
            padding: 12px 14px;
            text-align: left;
            border-bottom: 1px solid #e2e8f0;
        }

        th { background-color: #f1f5f9; color: var(--text-muted); font-weight: 600; }

        .badge {
            padding: 4px 10px;
            border-radius: 12px;
            font-size: 0.75rem;
            font-weight: 700;
            display: inline-block;
        }

        .badge-alto { background: #dcfce7; color: #15803d; }
        .badge-medio { background: #fef3c7; color: #b45309; }
        .badge-bajo { background: #fee2e2; color: #b91c1c; }
    </style>
</head>
<body>

    <header>
        <div>
            <h1>Institución Educativa La Esperanza</h1>
            <p>Dashboard de Resultados Saber — Simulacro #1 Diagnóstico (Grados 09° y 10°)</p>
        </div>
        <button class="btn-primary" onclick="exportExcel()">Exportar Excel</button>
    </header>

    <div class="controls">
        <select id="filterGrado" onchange="renderDashboard()">
            <option value="ALL">Todos los Grados (09° y 10°)</option>
            <option value="10°">Grado 10°</option>
            <option value="09°">Grado 09°</option>
        </select>
    </div>

    <div class="metrics-grid">
        <div class="card">
            <div class="metric-title">Promedio Global</div>
            <div class="metric-value" id="valPromedio">167.43 pts</div>
            <div class="metric-sub" id="valDE">Desviación Estándar: 24.03 pts</div>
        </div>
        <div class="card">
            <div class="metric-title">Mejor Área</div>
            <div class="metric-value" style="color: #059669;" id="valMejorArea">Lectura Crítica</div>
            <div class="metric-sub" id="valMejorScore">43.37 pts prom.</div>
        </div>
        <div class="card">
            <div class="metric-title">Área a Reforzar</div>
            <div class="metric-value" style="color: var(--danger);" id="valPeorArea">Matemáticas</div>
            <div class="metric-sub" id="valPeorScore">26.09 pts prom.</div>
        </div>
        <div class="card">
            <div class="metric-title">Total Estudiantes</div>
            <div class="metric-value" id="valTotalEst">23</div>
            <div class="metric-sub" id="valSubEst">13 de 09° | 10 de 10°</div>
        </div>
    </div>

    <!-- Sección DAFO -->
    <h2 style="margin-bottom: 12px; font-size: 1.2rem; color: var(--primary);">Matriz de Diagnóstico DAFO</h2>
    <div class="dafo-grid">
        <div class="dafo-card dafo-f">
            <h3 style="color: #15803d;">Fortalezas</h3>
            <ul>
                <li>Lectura Crítica se destaca como el área de mejor rendimiento promedio (43.37 pts).</li>
                <li>Liderazgo en ambos grados: Jhon Torres (10° - 213.67 pts) y Samuel Pinedo (09° - 203.38 pts).</li>
                <li>Desempeño sobresaliente individual en Inglés (María Fernanda Flores - 62.0 pts).</li>
            </ul>
        </div>
        <div class="dafo-card dafo-d">
            <h3 style="color: #b91c1c;">Debilidades</h3>
            <ul>
                <li>Matemáticas presenta el promedio más bajo del diagnóstico (26.09 pts).</li>
                <li>86.95% de los estudiantes (20 de 23) se ubicó en Nivel Bajo (< 200 pts).</li>
                <li>Ningún estudiante alcanzó el Nivel Alto (> 250 pts) en esta prueba inicial.</li>
            </ul>
        </div>
        <div class="dafo-card dafo-o">
            <h3 style="color: #1d4ed8;">Oportunidades</h3>
            <ul>
                <li>Sharik Castrillo (199.52 pts) y Valentina Sánchez (196.23 pts) están a un paso del Nivel Medio/Alto.</li>
                <li>Amplio margen de crecimiento conceptual en Matemáticas y Sociales antes del examen oficial.</li>
                <li>Implementación de círculos de estudio guiados por los estudiantes destacados.</li>
            </ul>
        </div>
        <div class="dafo-card dafo-a">
            <h3 style="color: #b45309;">Amenazas</h3>
            <ul>
                <li>Dificultades persistentes en competencia matemática que condicionan el Puntaje Global.</li>
                <li>Dispersión de resultados (DE de 24.03 pts) que evidencia brechas de aprendizaje heterogéneas.</li>
            </ul>
        </div>
    </div>

    <div class="charts-grid">
        <div class="chart-card">
            <h2>Promedios por Área</h2>
            <canvas id="barChart" height="140"></canvas>
        </div>
        <div class="chart-card">
            <h2>Distribución por Niveles de Desempeño</h2>
            <canvas id="pieChart"></canvas>
        </div>
    </div>

    <div class="chart-card">
        <h2>Ranking y Clasificación de Estudiantes</h2>
        <table id="studentsTable">
            <thead>
                <tr>
                    <th>Pos.</th>
                    <th>Grado</th>
                    <th>Nombre y Apellidos</th>
                    <th>Lectura</th>
                    <th>Matemáticas</th>
                    <th>Sociales</th>
                    <th>Naturales</th>
                    <th>Inglés</th>
                    <th>Puntaje Global</th>
                    <th>Nivel</th>
                </tr>
            </thead>
            <tbody></tbody>
        </table>
    </div>

    <script>
        const rawData = [
            {grado: '10°', nombre: 'TORRES RODRIGUEZ JHON ANDERSON', lec: 48.8, mat: 36.0, soc: 38.0, nat: 52.05, ing: 31.0, promedio: 41.17, global: 213.6730769230769},
            {grado: '10°', nombre: 'FLORES GUERRERO MARIA FERNANDA', lec: 36.6, mat: 26.0, soc: 46.0, nat: 44.60, ing: 62.0, promedio: 43.04, global: 200.6153846153846},
            {grado: '10°', nombre: 'CASTRILLO MARRIAGA SHARIK JHOANA', lec: 56.1, mat: 36.0, soc: 36.0, nat: 32.15, ing: 38.0, promedio: 39.65, global: 199.5192307692308},
            {grado: '10°', nombre: 'GUERRERO VELASQUEZ ANDRES', lec: 56.1, mat: 34.0, soc: 32.0, nat: 28.60, ing: 18.0, promedio: 33.74, global: 180.8076923076923},
            {grado: '10°', nombre: 'PAYARES HERNANDEZ LEISI CAROLINA', lec: 39.0, mat: 36.0, soc: 36.0, nat: 28.70, ing: 40.0, promedio: 35.94, global: 176.57692307692307},
            {grado: '10°', nombre: 'GASPAR SIERRA EDUARDO EMILIO', lec: 39.0, mat: 22.0, soc: 34.0, nat: 44.90, ing: 33.0, promedio: 34.58, global: 174.1153846153846},
            {grado: '10°', nombre: 'CASTRILLO MARRIAGA SHANI PATRICIA', lec: 53.7, mat: 16.0, soc: 30.0, nat: 36.00, ing: 29.0, promedio: 32.94, global: 167.73076923076923},
            {grado: '10°', nombre: 'GONZALEZ SOLIS JUAN CAMILO', lec: 41.5, mat: 32.0, soc: 24.0, nat: 32.30, ing: 38.0, promedio: 33.56, global: 164.3846153846154},
            {grado: '10°', nombre: 'VEGA CONTRERAS JAIDER DAVID', lec: 48.8, mat: 22.0, soc: 26.0, nat: 28.45, ing: 47.0, promedio: 34.45, global: 162.59615384615384},
            {grado: '10°', nombre: 'ORTIZ TALAIGUA MARIA ALEJANDRA', lec: 24.4, mat: 26.0, soc: 32.0, nat: 25.15, ing: 24.0, promedio: 26.31, global: 133.32692307692307},
            {grado: '09°', nombre: 'PINEDO BARRIOS SAMUEL', lec: 56.1, mat: 36.0, soc: 26.0, nat: 51.50, ing: 20.0, promedio: 37.92, global: 203.3846153846154},
            {grado: '09°', nombre: 'SANCHEZ FUENTES VALENTINA', lec: 51.1, mat: 38.0, soc: 42.0, nat: 30.30, ing: 26.0, promedio: 37.48, global: 196.23076923076923},
            {grado: '09°', nombre: 'BENITEZ URANGO DINA LUZ', lec: 51.2, mat: 30.0, soc: 26.0, nat: 41.30, ing: 35.0, promedio: 36.70, global: 184.80769230769232},
            {grado: '09°', nombre: 'ARIAS MONTALVO MARIA JOSE', lec: 41.5, mat: 22.0, soc: 26.0, nat: 48.65, ing: 40.0, promedio: 35.63, global: 174.78846153846155},
            {grado: '09°', nombre: 'VARGAS APARICIO ORLANDO MIGUEL', lec: 58.5, mat: 20.0, soc: 24.0, nat: 25.00, ing: 31.0, promedio: 31.70, global: 159.03846153846155},
            {grado: '09°', nombre: 'ACOSTA ZUÑIGA LUIS ENRIQUE', lec: 36.6, mat: 22.0, soc: 34.0, nat: 28.35, ing: 47.0, promedio: 33.59, global: 157.6346153846154},
            {grado: '09°', nombre: 'PEREZ VERGARA MARIA ANGEL', lec: 51.2, mat: 22.0, soc: 26.0, nat: 23.30, ing: 20.0, promedio: 28.50, global: 149.03846153846155},
            {grado: '09°', nombre: 'SUAREZ TALAIGUA MIGUEL ANDRES', lec: 39.0, mat: 18.0, soc: 34.0, nat: 23.15, ing: 40.0, promedio: 30.83, global: 147.09615384615384},
            {grado: '09°', nombre: 'HENAO SOLAR HORIANA ISABEL', lec: 29.3, mat: 22.0, soc: 22.0, nat: 42.90, ing: 26.0, promedio: 28.44, global: 144.07692307692307},
            {grado: '09°', nombre: 'MANCHEGO CESPEDES DARWIN ANDRES', lec: 36.6, mat: 22.0, soc: 28.0, nat: 28.85, ing: 26.0, promedio: 28.29, global: 143.21153846153845},
            {grado: '09°', nombre: 'TRUJILLO BERROCAL SAMUEL', lec: 43.9, mat: 20.0, soc: 26.0, nat: 28.70, ing: 16.0, promedio: 26.92, global: 143.0},
            {grado: '09°', nombre: 'FABRA MACEA ISAAC ELIAS', lec: 22.0, mat: 24.0, soc: 40.0, nat: 26.10, ing: 33.0, promedio: 29.02, global: 142.03846153846155},
            {grado: '09°', nombre: 'PINEDA OVIEDO NEIDIS YAMILES', lec: 36.6, mat: 18.0, soc: 24.0, nat: 26.60, ing: 31.0, promedio: 27.24, global: 133.30769230769232}
        ];

        let barChartInstance = null;
        let pieChartInstance = null;

        function renderDashboard() {
            const filter = document.getElementById('filterGrado').value;
            const filteredData = filter === 'ALL' ? rawData : rawData.filter(d => d.grado === filter);
            
            // Ordenar de mayor a menor por Puntaje Global sin alterar valores decimales
            filteredData.sort((a,b) => b.global - a.global);

            const total = filteredData.length;
            const cnt09 = filteredData.filter(d => d.grado === '09°').length;
            const cnt10 = filteredData.filter(d => d.grado === '10°').length;

            if (filter === 'ALL') {
                document.getElementById('valSubEst').innerText = `${cnt09} de 09° | ${cnt10} de 10°`;
            } else {
                document.getElementById('valSubEst').innerText = `Estudiantes evaluados de Grado ${filter}`;
            }

            const sumGlobal = filteredData.reduce((acc, d) => acc + d.global, 0);
            const avgGlobal = (sumGlobal / total).toFixed(2);
            
            // Desviación Estándar exacta
            const stdDev = Math.sqrt(filteredData.reduce((acc, d) => acc + Math.pow(d.global - parseFloat(avgGlobal), 2), 0) / total).toFixed(2);

            document.getElementById('valPromedio').innerText = avgGlobal + ' pts';
            document.getElementById('valDE').innerText = 'Desviación Estándar: ' + stdDev + ' pts';
            document.getElementById('valTotalEst').innerText = total;

            // Promedios por Área con sus décimas exactas
            const areaAvg = {
                'Lectura Crítica': (filteredData.reduce((a, b) => a + b.lec, 0) / total).toFixed(2),
                'Matemáticas': (filteredData.reduce((a, b) => a + b.mat, 0) / total).toFixed(2),
                'Sociales': (filteredData.reduce((a, b) => a + b.soc, 0) / total).toFixed(2),
                'Naturales': (filteredData.reduce((a, b) => a + b.nat, 0) / total).toFixed(2),
                'Inglés': (filteredData.reduce((a, b) => a + b.ing, 0) / total).toFixed(2)
            };

            const sortedAreas = Object.entries(areaAvg).sort((a,b) => b[1] - a[1]);
            document.getElementById('valMejorArea').innerText = sortedAreas[0][0];
            document.getElementById('valMejorScore').innerText = sortedAreas[0][1] + ' pts prom.';
            document.getElementById('valPeorArea').innerText = sortedAreas[sortedAreas.length - 1][0];
            document.getElementById('valPeorScore').innerText = sortedAreas[sortedAreas.length - 1][1] + ' pts prom.';

            // Niveles según límites exactos:
            // Bajo: < 200 | Medio: 200 - 250 | Alto: > 250
            let alto = 0, medio = 0, bajo = 0;
            filteredData.forEach(d => {
                const g = d.global;
                if (g > 250) alto++;
                else if (g >= 200) medio++;
                else bajo++;
            });

            // Actualizar Tabla con valores exactos
            const tbody = document.querySelector('#studentsTable tbody');
            tbody.innerHTML = '';
            filteredData.forEach((st, idx) => {
                const globalVal = st.global.toFixed(2);
                let badgeClass = 'badge-bajo';
                let levelText = 'Bajo (< 200)';

                if (st.global > 250) {
                    badgeClass = 'badge-alto';
                    levelText = 'Alto (> 250)';
                } else if (st.global >= 200) {
                    badgeClass = 'badge-medio';
                    levelText = 'Medio (200 - 250)';
                }

                tbody.innerHTML += `<tr>
                    <td><b>${idx + 1}</b></td>
                    <td>${st.grado}</td>
                    <td>${st.nombre}</td>
                    <td>${st.lec.toFixed(1)}</td>
                    <td>${st.mat.toFixed(1)}</td>
                    <td>${st.soc.toFixed(1)}</td>
                    <td>${st.nat.toFixed(2)}</td>
                    <td>${st.ing.toFixed(1)}</td>
                    <td><b>${globalVal}</b></td>
                    <td><span class="badge ${badgeClass}">${levelText}</span></td>
                </tr>`;
            });

            // Actualizar Gráficos
            renderBarChart(Object.keys(areaAvg), Object.values(areaAvg));
            renderPieChart([alto, medio, bajo]);
        }

        function renderBarChart(labels, values) {
            const ctx = document.getElementById('barChart').getContext('2d');
            if (barChartInstance) barChartInstance.destroy();

            barChartInstance = new Chart(ctx, {
                type: 'bar',
                data: {
                    labels: labels,
                    datasets: [{
                        label: 'Promedio de Puntaje',
                        data: values,
                        // Un color diferente y distinto para cada área
                        backgroundColor: [
                            'rgba(54, 162, 235, 0.85)',  // Azul - Lectura Crítica
                            'rgba(255, 159, 64, 0.85)',   // Naranja - Matemáticas
                            'rgba(255, 99, 132, 0.85)',   // Rojo/Rosa - Sociales y Ciudadanas
                            'rgba(75, 192, 192, 0.85)',   // Verde - Ciencias Naturales
                            'rgba(153, 102, 255, 0.85)'   // Morado - Inglés
                        ],
                        borderColor: [
                            '#1d4ed8', '#c2410c', '#be123c', '#047857', '#6d28d9'
                        ],
                        borderWidth: 1.5,
                        borderRadius: 6
                    }]
                },
                options: {
                    responsive: true,
                    scales: { y: { beginAtZero: true, max: 100 } },
                    plugins: { legend: { display: false } }
                }
            });
        }

        function renderPieChart(dataCounts) {
            const ctx = document.getElementById('pieChart').getContext('2d');
            if (pieChartInstance) pieChartInstance.destroy();

            pieChartInstance = new Chart(ctx, {
                type: 'pie',
                data: {
                    labels: ['Alto (> 250)', 'Medio (200 - 250)', 'Bajo (< 200)'],
                    datasets: [{
                        data: dataCounts,
                        backgroundColor: ['#10b981', '#f59e0b', '#ef4444']
                    }]
                },
                options: { responsive: true }
            });
        }

        function exportExcel() {
            const filter = document.getElementById('filterGrado').value;
            const filteredData = filter === 'ALL' ? rawData : rawData.filter(d => d.grado === filter);
            
            const processedData = filteredData.map(d => ({
                'Grado': d.grado,
                'Nombre y Apellidos': d.nombre,
                'Lectura Crítica': d.lec,
                'Matemáticas': d.mat,
                'Sociales y Ciudadanas': d.soc,
                'Ciencias Naturales': d.nat,
                'Inglés': d.ing,
                'Puntaje Global': Number(d.global.toFixed(2)),
                'Nivel': d.global > 250 ? 'Alto (> 250)' : (d.global >= 200 ? 'Medio (200 - 250)' : 'Bajo (< 200)')
            }));

            const ws = XLSX.utils.json_to_sheet(processedData);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "Resultados");
            XLSX.writeFile(wb, "Resultados_Simulacro1_La_Esperanza.xlsx");
        }

        // Render inicial
        renderDashboard();
    </script>
</body>
</html>
