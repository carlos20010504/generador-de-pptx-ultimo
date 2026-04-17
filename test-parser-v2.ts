
import * as XLSX from 'xlsx';
import { parsePresentationFromWorkbook } from './utils/excel-parser.js';

/**
 * MOCK TEST SUITE para validar el Pipeline de Datos
 * Propósito: Asegurar que el mapeo de columnas (case-insensitive) y la estructura de gráficas
 * funcione antes de entregar al usuario.
 */

function createMockWorkbook() {
    const wb = XLSX.utils.book_new();

    // 1. Hoja "Opo Mejora " (con espacio al final, como en los logs)
    const opoMejoraData = [
        ["REPORTE SEMANAL", "", ""], // Row 0 (basura)
        ["Fecha: 2026", "", ""],     // Row 1 (basura)
        ["HALLAZGO / ÁREA", "OPORTUNIDAD DE MEJORA", "ESTADO", "%", "¿EL CONTROL EXISTE?", "¿ESTÁ DOCUMENTADO?", "¿SE ESTÁ EJECUTANDO?", "¿ES EFECTIVO?", "¿REDUCE EL RIESGO?"], // Row 2 (headers)
        ["Proceso Financiero", "Revisión de facturas", "COMPLETADO", "1.0", "SI", "SI", "SI", "SI", "SI"],
        ["Proceso IT", "Backup diario", "EN PROCESO", "0.5", "SI", "NO", "SI", "NO", "SI"],
    ];
    const wsOpo = XLSX.utils.aoa_to_sheet(opoMejoraData);
    XLSX.utils.book_append_sheet(wb, wsOpo, "Opo Mejora ");

    // 2. Hoja "Comisiones- Base"
    const baseData = [
        ["TITULO", ""],
        ["Id Comisión", "Solicitante", "Ciudad Destino", "Fecha Inicio", "Fecha Fin", "Valor Total Solicitado", "Estado"], // Row 1
        ["C001", "Carlos Pinzon", "Medellín", 45627, 45630, 2500000, "APROBADO LIDER"],
        ["C002", "Ana Maria", "Bogotá", 45631, 45635, 1800000, "PENDIENTE"],
    ];
    const wsBase = XLSX.utils.aoa_to_sheet(baseData);
    XLSX.utils.book_append_sheet(wb, wsBase, "Comisiones- Base");

    return wb;
}

async function runTest() {
    console.log("🚀 Iniciando pruebas de validación del Parser...");
    const wb = createMockWorkbook();

    try {
        const result = parsePresentationFromWorkbook(wb, { validate: true });
        
        console.log(`📊 Total diapositivas generadas: ${result.slides.length}`);

        // VALIDACIÓN: Oportunidades de Mejora
        const opo: any = result.slides.find((s: any) => s.title === "Oportunidades de Mejora");
        if (!opo) {
            console.error("❌ ERROR: No se encontró la slide de 'Oportunidades de Mejora'");
        } else {
            console.log("✅ Slide 'Opo Mejora' encontrada.");
            console.log(`   - Filas: ${opo.rows.length}`);
            console.log(`   - Gráficas: ${opo.charts?.length || 0}`);
            
            if (opo.rows.length > 0) {
                const first = opo.rows[0];
                console.log("   - Headers detectados en primera fila:", Object.keys(first));
            }
        }

        // VALIDACIÓN: Charts
        const hasCharts = result.slides.some((s: any) => s.charts && s.charts.length > 0);
        if (hasCharts) {
            console.log("✅ Se detectaron gráficas en el payload.");
        } else {
            console.error("❌ ERROR: No se generaron gráficas. Hay un problema en la estructura.");
        }

    } catch (e) {
        console.error("❌ Falla crítica en el test:", e);
    }
}

runTest();
