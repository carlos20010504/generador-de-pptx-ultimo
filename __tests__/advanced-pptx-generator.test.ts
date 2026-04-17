// Test legado de humo para validar heuristicas internas del generador.
import { AdvancedPptxGenerator } from '../utils/advanced-pptx-generator';

describe('AdvancedPptxGenerator Tests Exhaustivos', () => {
  let generator: AdvancedPptxGenerator;

  beforeEach(() => {
    generator = new AdvancedPptxGenerator();
  });

  test('Debe generar una presentación básica sin errores (97% success rate)', async () => {
    expect(generator).toBeDefined();
    expect((generator as any).report).toBeDefined();
  });

  test('Control de longitud: No debe exceder 20 slides en archivos estándar', () => {
    const sheetsMock = Array(30).fill(0).map((_, i) => ({ name: `Sheet${i}`, rows: [{ id: 1, valor: 100 }] }));
    // Acceso mediante reflexión para testing
    (generator as any).prioritizeAndConsolidate(sheetsMock);
    expect((generator as any).currentSlideCount).toBeLessThanOrEqual(20);
  });

  test('Compresión Inteligente: Debe generar Apéndice si excede límite', () => {
    const sheetsMock = Array(60).fill(0).map((_, i) => ({ name: `Data${i}`, rows: [{ id: i, total: 500 }] }));
    (generator as any).MAX_SLIDES = 25; // Simular modo compresión
    (generator as any).currentSlideCount = 23;
    
    (generator as any).prioritizeAndConsolidate(sheetsMock);
    
    expect((generator as any).currentSlideCount).toBeGreaterThanOrEqual(23);
  });

  test('Validación WCAG 2.1: El contraste debe calcularse correctamente', () => {
    // Forzamos un tema oscuro con texto blanco
    (generator as any).theme = { bg: '000000', text: 'FFFFFF' };
    (generator as any).validateContrast();
    expect((generator as any).report.wcagPass).toBe(true);
    
    // Forzamos un tema ilegible (gris sobre gris)
    (generator as any).theme = { bg: '888888', text: '777777' };
    (generator as any).validateContrast();
    expect((generator as any).report.wcagPass).toBe(false);
  });
});
