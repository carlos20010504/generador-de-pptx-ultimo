
import pptxgen from 'pptxgenjs';
import { buildSocyaPresentation, SocyaSlideJSON } from './utils/socya-renderer';
import fs from 'fs';

const mockSlides: SocyaSlideJSON[] = [
  {
    type: 'title',
    title: 'Reporte de Auditoría de Comisiones',
    subtitle: 'Análisis integral generado con Socya PPTX Engine',
    content: null
  },
  {
    type: 'kpi_row',
    title: 'Resumen Ejecutivo',
    content: [
      { label: 'Total Comisiones', value: '45' },
      { label: 'Valor Total', value: '$124.5M' },
      { label: 'Solicitantes', value: '18' },
      { label: 'Ciudades Destino', value: '12' },
      { label: 'Centros de Costos', value: '6' },
    ]
  },
  {
    type: 'chart',
    title: 'Distribución de Comisiones por Estado',
    content: {
      name: 'Estados',
      labels: ['Contabilizado', 'Legalizado', 'Rechazado', 'Solicitado'],
      values: [20, 15, 5, 5],
      barDir: 'bar'
    }
  },
  {
    type: 'table',
    title: 'Muestra de Comisiones',
    subtitle: 'Registros 1 - 10 de 120',
    content: {
      headers: ['ID', 'Solicitante', 'Destino', 'Valor Total', 'Estado'],
      rows: [
        ['COM-001', 'Juan Perez', 'Bogotá', '$1.200.000', 'CONTABILIZADO'],
        ['COM-002', 'Maria Lopez', 'Medellín', '$850.000', 'LEGALIZADO'],
        ['COM-003', 'Carlos Ruiz', 'Cali', '$2.100.000', 'RECHAZADO'],
        ['COM-004', 'Ana Gomez', 'Barranquilla', '$450.000', 'SOLICITADO'],
        ['COM-005', 'Pedro Sanchez', 'Cartagena', '$1.100.000', 'CONTABILIZADO'],
      ]
    },
    detail_link: 'https://example.com'
  },
  {
    type: 'text_bullets',
    title: 'Resumen de Hallazgos Clave',
    content: [
      'Se identificaron duplicidades en el cobro de viáticos de transporte.',
      'Falta de soportes en el 15% de las legalizaciones de alimentación.',
      'Retrasos en la aprobación de comisiones por parte de los líderes regionales.',
      'Oportunidad de optimización en la reserva de tiquetes aéreos con 3 días de antelación.'
    ]
  },
  {
    type: 'closing',
    title: 'Fin del Reporte',
    subtitle: '¡Gracias por su atención!',
    content: null
  }
];

const prs = new pptxgen();
prs.layout = 'LAYOUT_WIDE';
buildSocyaPresentation(prs, mockSlides);

prs.writeFile({ fileName: 'test_output.pptx' })
  .then(() => console.log('Presentation saved as test_output.pptx'))
  .catch(err => console.error(err));
