/* Ficha Social UNCP — entrevista completa alineada a Ficha social JULIO 2026.xlsx */
/* global ExcelJS, PLANTILLA_FICHA_B64 */

const SECCIONES = {
  1: 'I. Datos Generales del Socio(a)',
  2: 'II. Caracterización Familiar',
  3: 'III. Composición del Núcleo Familiar',
  4: 'IV. Apreciación del Núcleo Familiar',
  5: 'V. Condición de Discapacidad Visual',
  6: 'VI. Estado de Salud',
  7: 'VII. Situación Laboral y Económica',
  8: 'VIII. Condición de Vivienda',
  9: 'Diagnóstico, Mortuoria y Registro'
};

const PREGUNTAS = [
  // —— I. Datos generales ——
  {
    id: 'nombres', seccion: 1, required: true, tipo: 'texto',
    pregunta: '¿Cuáles son los apellidos y nombres completos del socio o socia?',
    ayuda: 'Apellidos y luego los nombres.',
    vozPregunta: 'Pregunta: Apellidos y nombres completos del socio o socia.'
  },
  {
    id: 'dni', seccion: 1, tipo: 'texto', inputmode: 'numeric', maxlength: 12,
    pregunta: '¿Cuál es el número de DNI o carné?',
    ayuda: 'Solo números, máximo 12 dígitos.',
    vozPregunta: 'Pregunta: Número de DNI o carné.'
  },
  {
    id: 'fecha_nac', seccion: 1, tipo: 'fecha',
    pregunta: '¿Cuál es la fecha de nacimiento?',
    vozPregunta: 'Pregunta: Fecha de nacimiento.'
  },
  {
    id: 'edad', seccion: 1, tipo: 'numero', min: 0, max: 120,
    pregunta: '¿Qué edad tiene, en años?',
    vozPregunta: 'Pregunta: Edad en años.'
  },
  {
    id: 'sexo', seccion: 1, tipo: 'radio',
    pregunta: '¿Cuál es el sexo?',
    opciones: [
      { value: 'M', label: 'Masculino (M)' },
      { value: 'F', label: 'Femenino (F)' }
    ],
    vozPregunta: 'Pregunta: Sexo. Masculino o femenino.'
  },
  {
    id: 'depto_nac', seccion: 1, tipo: 'texto',
    pregunta: 'Lugar de nacimiento: ¿en qué departamento nació?',
    vozPregunta: 'Departamento de nacimiento.'
  },
  {
    id: 'prov_nac', seccion: 1, tipo: 'texto',
    pregunta: '¿En qué provincia nació?',
    vozPregunta: 'Provincia de nacimiento.'
  },
  {
    id: 'dist_nac', seccion: 1, tipo: 'texto',
    pregunta: '¿En qué distrito nació?',
    vozPregunta: 'Distrito de nacimiento.'
  },
  {
    id: 'estado_civil', seccion: 1, tipo: 'radio',
    pregunta: '¿Cuál es el estado civil?',
    opciones: [
      { value: 'soltero', label: 'Soltero(a)' },
      { value: 'casado', label: 'Casado(a)' },
      { value: 'viudo', label: 'Viudo(a)' },
      { value: 'conviviente', label: 'Conviviente' },
      { value: 'divorciado', label: 'Divorciado(a)' },
      { value: 'separado', label: 'Separado(a)' }
    ],
    vozPregunta: 'Estado civil: soltero, casado, viudo, conviviente, divorciado o separado.'
  },
  {
    id: 'direccion', seccion: 1, tipo: 'texto',
    pregunta: '¿Cuál es la dirección actual?',
    ayuda: 'Calle, número, urbanización o asentamiento.',
    vozPregunta: 'Dirección actual.'
  },
  {
    id: 'referencia', seccion: 1, tipo: 'texto',
    pregunta: '¿Qué referencia tiene la dirección?',
    ayuda: 'Ejemplo: cerca de la plaza, frente al mercado.',
    vozPregunta: 'Referencia de la dirección.'
  },
  {
    id: 'instruccion', seccion: 1, tipo: 'select',
    pregunta: '¿Cuál es el nivel educativo o grado de instrucción?',
    opciones: [
      { value: 'sin-estudios', label: 'Sin estudios' },
      { value: 'primaria', label: 'Primaria incompleta' },
      { value: 'primaria-completa', label: 'Primaria completa' },
      { value: 'secundaria', label: 'Secundaria incompleta' },
      { value: 'secundaria-completa', label: 'Secundaria completa' },
      { value: 'tecnico', label: 'Técnico' },
      { value: 'universitario', label: 'Universitario' }
    ],
    vozPregunta: 'Nivel educativo. Por ejemplo: primaria, secundaria, técnico o universitario.'
  },
  {
    id: 'telefono', seccion: 1, tipo: 'tel',
    pregunta: '¿Cuál es el celular de contacto?',
    vozPregunta: 'Celular de contacto.'
  },
  {
    id: 'correo', seccion: 1, tipo: 'texto',
    pregunta: '¿Cuál es el correo electrónico? (puede omitir si no tiene)',
    vozPregunta: 'Correo electrónico. Puede omitir si no tiene.'
  },
  {
    id: 'c1_nombre', seccion: 1, tipo: 'texto',
    pregunta: 'Contacto de emergencia 1: ¿nombre completo?',
    vozPregunta: 'Nombre del primer contacto de emergencia.'
  },
  {
    id: 'c1_par', seccion: 1, tipo: 'texto',
    pregunta: 'Contacto 1: ¿parentesco?',
    ayuda: 'Ejemplo: hijo, esposa, hermano.',
    vozPregunta: 'Parentesco del primer contacto.'
  },
  {
    id: 'c1_cel', seccion: 1, tipo: 'tel',
    pregunta: 'Contacto 1: ¿celular?',
    vozPregunta: 'Celular del primer contacto de emergencia.'
  },
  {
    id: 'c2_nombre', seccion: 1, tipo: 'texto',
    pregunta: 'Contacto de emergencia 2: ¿nombre completo? (puede omitir)',
    vozPregunta: 'Nombre del segundo contacto. Puede omitir si no hay.'
  },
  {
    id: 'c2_par', seccion: 1, tipo: 'texto',
    pregunta: 'Contacto 2: ¿parentesco?',
    vozPregunta: 'Parentesco del segundo contacto.'
  },
  {
    id: 'c2_cel', seccion: 1, tipo: 'tel',
    pregunta: 'Contacto 2: ¿celular?',
    vozPregunta: 'Celular del segundo contacto.'
  },

  // —— II. Caracterización familiar ——
  {
    id: 'es_padre_madre', seccion: 2, tipo: 'radio',
    pregunta: '¿Usted es padre o madre?',
    opciones: [
      { value: 'si', label: 'Sí' },
      { value: 'no', label: 'No' }
    ],
    vozPregunta: '¿Usted es padre o madre? Sí o no.'
  },
  {
    id: 'n_hijos', seccion: 2, tipo: 'numero', min: 0, max: 30,
    pregunta: '¿Cuántos hijos tiene? (0 si no tiene)',
    vozPregunta: 'Número de hijos.'
  },
  {
    id: 'n_hijos_disc', seccion: 2, tipo: 'numero', min: 0, max: 30,
    pregunta: '¿Cuántos hijos tienen discapacidad visual?',
    vozPregunta: 'Número de hijos con discapacidad visual.'
  },
  {
    id: 'otros_fam_disc', seccion: 2, tipo: 'texto',
    pregunta: '¿Tiene otros familiares con discapacidad visual? Mencione parentesco o diga “ninguno”.',
    vozPregunta: 'Otros familiares con discapacidad visual. Mencione parentesco o diga ninguno.'
  },

  // —— III. Composición del núcleo ——
  // Tras num_familia, construirColaBase() inserta por cada miembro (1–7):
  // nombres, parentesco, edad, nivel educativo, ocupación, aporta ingresos, disc./enfermedad.
  // Luego tipo de familia (mismo orden que la ficha oficial).
  {
    id: 'num_familia', seccion: 3, tipo: 'numero', min: 0, max: 7,
    pregunta: '¿Cuántas personas viven con el socio, sin contarlo a él o ella? (0 a 7)',
    ayuda: 'Por cada persona se preguntará: apellidos y nombres, parentesco, edad, nivel educativo, ocupación, si aporta ingresos y discapacidad o enfermedad. Después se pregunta el tipo de familia.',
    vozPregunta: 'Composición del núcleo familiar. ¿Cuántas personas viven con el socio, sin contarlo? Del 0 al 7. Luego preguntaré los datos de cada una: nombres, parentesco, edad, nivel educativo, ocupación, si aporta ingresos y si tiene discapacidad o enfermedad.'
  },
  {
    id: 'tipo_familia', seccion: 3, tipo: 'radio',
    pregunta: '¿Qué tipo de familia es?',
    ayuda: 'Después de registrar a cada miembro del núcleo.',
    opciones: [
      { value: 'unipersonal', label: 'Unipersonal' },
      { value: 'nuclear', label: 'Nuclear' },
      { value: 'monoparental', label: 'Monoparental' },
      { value: 'extensa', label: 'Extensa' },
      { value: 'otros', label: 'Otros' }
    ],
    vozPregunta: 'Tipo de familia: unipersonal, nuclear, monoparental, extensa u otros.'
  },
  {
    id: 'tipo_familia_otro', seccion: 3, tipo: 'texto',
    pregunta: 'Si eligió “Otros” en tipo de familia, especifique (u omita)',
    vozPregunta: 'Especifique otro tipo de familia. Puede omitir.'
  },

  // —— IV. Apreciación ——
  {
    id: 'relaciones_fam', seccion: 4, tipo: 'radio',
    pregunta: '¿Cómo califica las relaciones familiares?',
    opciones: [
      { value: 'buena', label: 'Buena' },
      { value: 'regular', label: 'Regular' },
      { value: 'mala', label: 'Mala' }
    ],
    vozPregunta: 'Relaciones familiares: buena, regular o mala.'
  },
  {
    id: 'comunicacion_fam', seccion: 4, tipo: 'radio',
    pregunta: '¿Cómo califica la comunicación familiar?',
    opciones: [
      { value: 'buena', label: 'Buena' },
      { value: 'regular', label: 'Regular' },
      { value: 'mala', label: 'Mala' }
    ],
    vozPregunta: 'Comunicación familiar: buena, regular o mala.'
  },
  {
    id: 'red_apoyo', seccion: 4, tipo: 'texto',
    pregunta: '¿Quiénes conforman su principal red de apoyo?',
    vozPregunta: 'Principal red de apoyo. Por ejemplo familia, amigos, vecinos, institución.'
  },
  {
    id: 'apreciacion_prof', seccion: 4, tipo: 'textarea',
    pregunta: 'Apreciación profesional (para quien entrevista)',
    ayuda: 'Observaciones del profesional o acompañante sobre el núcleo familiar.',
    vozPregunta: 'Apreciación profesional sobre el núcleo familiar.'
  },

  // —— V. Discapacidad visual ——
  {
    id: 'tipo_discapacidad', seccion: 5, tipo: 'radio',
    pregunta: '¿Cuál es la condición de discapacidad visual?',
    opciones: [
      { value: 'ceguera', label: 'Ceguera total' },
      { value: 'baja-vision', label: 'Baja visión' }
    ],
    vozPregunta: 'Condición: ceguera total o baja visión.'
  },
  {
    id: 'condicion', seccion: 5, tipo: 'radio',
    pregunta: '¿El origen es congénito o adquirido?',
    opciones: [
      { value: 'congenita', label: 'Congénita' },
      { value: 'adquirida', label: 'Adquirida' }
    ],
    vozPregunta: 'Origen: congénita o adquirida.'
  },
  {
    id: 'edad_adquisicion', seccion: 5, tipo: 'numero', min: 0, max: 120,
    pregunta: '¿A qué edad adquirió la discapacidad? (0 si es congénita)',
    vozPregunta: 'Edad en que adquirió la discapacidad. Cero si es congénita.'
  },
  {
    id: 'certificado_disc', seccion: 5, tipo: 'radio',
    pregunta: '¿Cuenta con certificado de discapacidad?',
    opciones: [
      { value: 'si', label: 'Sí' },
      { value: 'no', label: 'No' }
    ],
    vozPregunta: '¿Cuenta con certificado de discapacidad? Sí o no.'
  },
  {
    id: 'carnet_conadis', seccion: 5, tipo: 'radio',
    pregunta: '¿Cuenta con carnet de CONADIS?',
    opciones: [
      { value: 'si', label: 'Sí' },
      { value: 'no', label: 'No' }
    ],
    vozPregunta: '¿Cuenta con carnet de CONADIS? Sí o no.'
  },
  {
    id: 'atencion_oftal', seccion: 5, tipo: 'radio',
    pregunta: '¿Recibe atención oftalmológica actualmente?',
    opciones: [
      { value: 'si', label: 'Sí' },
      { value: 'no', label: 'No' }
    ],
    vozPregunta: '¿Recibe atención oftalmológica actualmente? Sí o no.'
  },
  {
    id: 'lugar_atencion', seccion: 5, tipo: 'texto',
    pregunta: '¿En qué lugar o institución recibe la atención oftalmológica? (omitir si no aplica)',
    vozPregunta: 'Lugar o institución de atención oftalmológica.'
  },

  // —— VI. Salud ——
  {
    id: 'seguro', seccion: 6, tipo: 'radio',
    pregunta: '¿Qué seguro de salud tiene?',
    opciones: [
      { value: 'sis', label: 'SIS' },
      { value: 'essalud', label: 'EsSalud' },
      { value: 'privado', label: 'Privado' },
      { value: 'otros', label: 'Otros' }
    ],
    vozPregunta: 'Seguro de salud: SIS, EsSalud, privado u otros.'
  },
  {
    id: 'seguro_otro', seccion: 6, tipo: 'texto',
    pregunta: 'Si eligió “Otros” en seguro, especifique (u omita)',
    vozPregunta: 'Especifique el otro seguro. Puede omitir.'
  },
  {
    id: 'enfermedades_cronicas', seccion: 6, tipo: 'textarea',
    pregunta: '¿Presenta enfermedades crónicas? Especifique o diga “ninguna”.',
    vozPregunta: 'Enfermedades crónicas. Especifique o diga ninguna.'
  },
  {
    id: 'autonomia', seccion: 6, tipo: 'radio',
    pregunta: 'Autonomía funcional y movilidad',
    opciones: [
      { value: 'independiente', label: 'Independiente' },
      { value: 'parcial', label: 'Requiere apoyo parcial' },
      { value: 'permanente', label: 'Requiere apoyo permanente' }
    ],
    vozPregunta: 'Autonomía: independiente, apoyo parcial o apoyo permanente.'
  },

  // —— VII. Laboral ——
  {
    id: 'ingreso', seccion: 7, tipo: 'radio',
    pregunta: '¿Cuál es el ingreso mensual aproximado del socio(a)?',
    opciones: [
      { value: 'sin-ingresos', label: 'Sin ingresos' },
      { value: 'menos-500', label: 'Menos de S/ 500' },
      { value: '500-900', label: 'Entre S/ 500 y S/ 900' },
      { value: '900-1500', label: 'Entre S/ 900 y S/ 1500' },
      { value: 'mas-1500', label: 'Más de S/ 1500' }
    ],
    vozPregunta: 'Ingreso mensual: sin ingresos, menos de 500, entre 500 y 900, entre 900 y 1500, o más de 1500 soles.'
  },
  {
    id: 'laborando', seccion: 7, tipo: 'radio',
    pregunta: '¿Labora actualmente?',
    opciones: [
      { value: 'si', label: 'Sí' },
      { value: 'no', label: 'No' }
    ],
    vozPregunta: '¿Labora actualmente? Sí o no.'
  },
  {
    id: 'ocupacion', seccion: 7, tipo: 'texto',
    pregunta: '¿Cuál es su ocupación actual? (omitir si no labora)',
    vozPregunta: 'Ocupación actual.'
  },
  {
    id: 'profesion', seccion: 7, tipo: 'texto',
    pregunta: '¿Cuál es su profesión u oficio?',
    vozPregunta: 'Profesión u oficio.'
  },
  {
    id: 'fuente_ingresos', seccion: 7, tipo: 'radio',
    pregunta: '¿Cuál es la fuente principal de ingresos del socio(a)?',
    opciones: [
      { value: 'dependiente', label: 'Trabajo dependiente' },
      { value: 'independiente', label: 'Trabajo independiente' },
      { value: 'pension', label: 'Pensión / jubilación' },
      { value: 'programa', label: 'Programa social' },
      { value: 'apoyo-familiar', label: 'Apoyo familiar' },
      { value: 'otros', label: 'Otros' }
    ],
    vozPregunta: 'Fuente principal de ingresos: trabajo dependiente, independiente, pensión, programa social, apoyo familiar u otros.'
  },
  {
    id: 'programa_social_detalle', seccion: 7, tipo: 'texto',
    pregunta: 'Si tiene programa social u “otros” ingresos, especifique (u omita)',
    ayuda: 'Ejemplo: Pensión 65, Contigo, CONADIS.',
    vozPregunta: 'Detalle de programa social u otros. Puede omitir.'
  },
  {
    id: 'otros_ingresos_hogar', seccion: 7, tipo: 'radio',
    pregunta: '¿Cuenta con otros ingresos en el hogar?',
    opciones: [
      { value: 'si', label: 'Sí' },
      { value: 'no', label: 'No' }
    ],
    vozPregunta: '¿Hay otros ingresos en el hogar? Sí o no.'
  },
  {
    id: 'principal_aportante', seccion: 7, tipo: 'texto',
    pregunta: '¿Quién es el principal aportante económico en el hogar?',
    vozPregunta: 'Principal aportante económico en el hogar.'
  },
  {
    id: 'sisfoh', seccion: 7, tipo: 'radio',
    pregunta: 'Clasificación socioeconómica SISFOH',
    opciones: [
      { value: 'no-tiene', label: 'No tiene' },
      { value: 'no-sabe', label: 'No sabe' },
      { value: 'pobre', label: 'Pobre' },
      { value: 'pobre-extremo', label: 'Pobre extremo' },
      { value: 'no-pobre', label: 'No pobre' }
    ],
    vozPregunta: 'Clasificación SISFOH: no tiene, no sabe, pobre, pobre extremo o no pobre.'
  },

  // —— VIII. Vivienda ——
  {
    id: 'tenencia', seccion: 8, tipo: 'radio',
    pregunta: 'Tenencia de la vivienda',
    opciones: [
      { value: 'propia', label: 'Propia' },
      { value: 'alquilada', label: 'Alquilada' },
      { value: 'cedida', label: 'Cedida' },
      { value: 'otro', label: 'Otro' }
    ],
    vozPregunta: 'Tenencia: propia, alquilada, cedida u otro.'
  },
  {
    id: 'tenencia_otro', seccion: 8, tipo: 'texto',
    pregunta: 'Si tenencia es “Otro”, especifique (u omita)',
    vozPregunta: 'Especifique otra tenencia. Puede omitir.'
  },
  {
    id: 'tipo_vivienda', seccion: 8, tipo: 'radio',
    pregunta: 'Tipo de vivienda',
    opciones: [
      { value: 'casa', label: 'Casa independiente' },
      { value: 'departamento', label: 'Departamento' },
      { value: 'cuarto', label: 'Cuarto' },
      { value: 'quinta', label: 'Quinta' },
      { value: 'multifamiliar', label: 'Multifamiliar' },
      { value: 'otro', label: 'Otro' }
    ],
    vozPregunta: 'Tipo de vivienda: casa independiente, departamento, cuarto, quinta, multifamiliar u otro.'
  },
  {
    id: 'tipo_vivienda_otro', seccion: 8, tipo: 'texto',
    pregunta: 'Si tipo de vivienda es “Otro”, especifique (u omita)',
    vozPregunta: 'Especifique otro tipo de vivienda.'
  },
  {
    id: 'n_ambientes', seccion: 8, tipo: 'numero', min: 0, max: 50,
    pregunta: '¿Número de ambientes de la vivienda utilizados? No incluir cocina,baños, pasadizos ni cochera ',
    vozPregunta: 'Número de ambientes.'
  },
  {
    id: 'distribucion_dorm', seccion: 8, tipo: 'radio',
    pregunta: 'Distribución del dormitorio del socio',
    opciones: [
      { value: 'propio', label: 'Dormitorio propio' },
      { value: 'comparte-dorm', label: 'Comparte dormitorio' },
      { value: 'comparte-hab', label: 'Comparte habitación' }
    ],
    vozPregunta: '¿Tiene dormitorio propio, comparte dormitorio o comparte habitación?'
  },
  {
    id: 'servicios', seccion: 8, tipo: 'checkbox',
    pregunta: '¿Con qué servicios básicos cuenta? (puede marcar varios)',
    opciones: [
      { value: 'agua', label: 'Agua potable' },
      { value: 'desague', label: 'Desagüe' },
      { value: 'luz', label: 'Energía eléctrica' },
      { value: 'internet', label: 'Internet' }
    ],
    vozPregunta: 'Servicios básicos: diga agua, desagüe, luz e internet según corresponda.'
  },
  {
    id: 'barreras', seccion: 8, tipo: 'radio',
    pregunta: 'Condiciones de accesibilidad en el hogar',
    opciones: [
      { value: 'si', label: 'Presenta barreras' },
      { value: 'no', label: 'No presenta barreras' }
    ],
    vozPregunta: '¿Presenta barreras de accesibilidad en el hogar o no presenta barreras?'
  },
  {
    id: 'barreras_especificar', seccion: 8, tipo: 'texto',
    pregunta: 'Si presenta barreras, especifique (u omita)',
    vozPregunta: 'Especifique las barreras de accesibilidad. Puede omitir.'
  },

  // —— Cierre ——
  {
    id: 'diagnostico_social', seccion: 9, tipo: 'textarea',
    pregunta: 'Diagnóstico social (profesional)',
    ayuda: 'Síntesis del asistente social o de quien realiza la entrevista.',
    vozPregunta: 'Diagnóstico social. Espacio para el profesional.'
  },
  {
    id: 'derecho_mortuoria', seccion: 9, tipo: 'textarea',
    pregunta: 'Información sobre derecho de mortuoria',
    vozPregunta: 'Derecho de mortuoria. Puede omitir si no aplica.'
  },
  {
    id: 'fecha_entrevista', seccion: 9, tipo: 'fecha',
    pregunta: '¿Cuál es la fecha de la entrevista?',
    vozPregunta: 'Fecha de la entrevista.'
  },
  {
    id: 'profesional', seccion: 9, tipo: 'texto',
    pregunta: '¿Nombre del profesional responsable?',
    vozPregunta: 'Nombre del profesional responsable.'
  },
  {
    id: 'firma', seccion: 9, tipo: 'textarea',
    pregunta: 'Firma o huella del socio (nota: firmó / dejó huella / no pudo firmar)',
    vozPregunta: 'Nota sobre firma o huella del socio.'
  }
];

/** Campos de la ficha por cada persona del núcleo (filas de la tabla III). */
const CAMPOS_MIEMBRO = ['nombre', 'parentesco', 'edad', 'instruccion', 'ocupacion', 'aporta', 'disc_enf'];
const MAX_MIEMBROS_FAMILIA = 7;

const ORDENAL_MIEMBRO = {
  1: 'primera', 2: 'segunda', 3: 'tercera', 4: 'cuarta',
  5: 'quinta', 6: 'sexta', 7: 'séptima'
};

function plantillaMiembro(n) {
  const ord = ORDENAL_MIEMBRO[n] || String(n);
  return [
    {
      id: `fam${n}_nombre`, seccion: 3, _miembro: n, tipo: 'texto',
      pregunta: `Familiar ${n}: apellidos y nombres`,
      ayuda: 'Tabla III — Apellidos y nombres de quien vive con el socio.',
      vozPregunta: `Datos de la ${ord} persona del núcleo familiar. Apellidos y nombres.`
    },
    {
      id: `fam${n}_parentesco`, seccion: 3, _miembro: n, tipo: 'texto',
      pregunta: `Familiar ${n}: parentesco con el socio`,
      ayuda: 'Ejemplo: cónyuge, hijo(a), padre, madre, hermano(a), nieto(a), otro.',
      vozPregunta: `Familiar ${n}. ¿Qué parentesco tiene con el socio?`
    },
    {
      id: `fam${n}_edad`, seccion: 3, _miembro: n, tipo: 'numero', min: 0, max: 120,
      pregunta: `Familiar ${n}: edad (años)`,
      vozPregunta: `Familiar ${n}. ¿Qué edad tiene, en años?`
    },
    {
      id: `fam${n}_instruccion`, seccion: 3, _miembro: n, tipo: 'texto',
      pregunta: `Familiar ${n}: nivel educativo`,
      ayuda: 'Ejemplo: sin estudios, primaria, secundaria, técnico, universitario.',
      vozPregunta: `Familiar ${n}. ¿Cuál es su nivel educativo?`
    },
    {
      id: `fam${n}_ocupacion`, seccion: 3, _miembro: n, tipo: 'texto',
      pregunta: `Familiar ${n}: ocupación`,
      ayuda: 'Trabajo o actividad principal. Si no trabaja, diga “ninguna” o “estudiante”.',
      vozPregunta: `Familiar ${n}. ¿Cuál es su ocupación?`
    },
    {
      id: `fam${n}_aporta`, seccion: 3, _miembro: n, tipo: 'radio',
      pregunta: `Familiar ${n}: ¿aporta ingresos al hogar?`,
      opciones: [
        { value: 'si', label: 'Sí' },
        { value: 'no', label: 'No' }
      ],
      vozPregunta: `Familiar ${n}. ¿Aporta ingresos al hogar? Sí o no.`
    },
    {
      id: `fam${n}_disc_enf`, seccion: 3, _miembro: n, tipo: 'texto',
      pregunta: `Familiar ${n}: discapacidad o enfermedad`,
      ayuda: 'Especifique o diga “ninguna”.',
      vozPregunta: `Familiar ${n}. ¿Tiene alguna discapacidad o enfermedad? Especifique o diga ninguna.`
    }
  ];
}

/** ¿Esta pregunta del núcleo aplica según la cantidad indicada? */
function esPreguntaActiva(q) {
  if (!q) return false;
  if (q._miembro != null) {
    const n = parseInt(respuestas.num_familia, 10) || 0;
    return q._miembro >= 1 && q._miembro <= n;
  }
  return true;
}

/** Busca el siguiente/anterior índice de pregunta activa (dir: +1 o -1). */
function buscarIndiceActivo(desde, dir) {
  let i = desde;
  while (i >= 0 && i < cola.length) {
    if (esPreguntaActiva(cola[i])) return i;
    i += dir;
  }
  return dir > 0 ? cola.length : -1;
}

function limpiarMiembrosFueraDeRango(cantidad) {
  for (let m = cantidad + 1; m <= MAX_MIEMBROS_FAMILIA; m++) {
    CAMPOS_MIEMBRO.forEach(campo => {
      delete respuestas[`fam${m}_${campo}`];
    });
  }
}

function aplicarCantidadFamilia(val) {
  let n = parseInt(val, 10);
  if (isNaN(n) || n < 0) n = 0;
  if (n > MAX_MIEMBROS_FAMILIA) n = MAX_MIEMBROS_FAMILIA;
  respuestas.num_familia = n;
  limpiarMiembrosFueraDeRango(n);
  miembrosInsertados = n;
  return n;
}

// ═══════════════════════════════════════════
// ESTADO
// ═══════════════════════════════════════════
let cola = [];
let indice = 0;
let respuestas = {};
let velocidadVoz = 0.9;
let reconocimiento = null;
let escuchando = false;
let miembrosInsertados = 0;
let hablarToken = 0;
let hablarTimerSeguridad = null;
let vocesListas = false;
let transcriptParcial = '';
let transcriptFinal = '';
let microStream = null; // permiso micrófono mantenido
let silencioTimer = null; // auto-detener mic tras silencio
let procesandoDictado = false; // evita doble procesamiento
let vozAutoAvanceTimer = null;
/** Una vez activado (permiso + 1.er clic o al comenzar), escucha solo tras cada pregunta. */
let modoVozContinua = false;
let escuchaAutoTimer = null;
let escuchaAutoToken = 0; // cancela programaciones viejas al cambiar de pregunta
/**
 * Lectura en voz alta (TTS / narrador).
 * Por defecto OFF: no suena el narrador hasta que el usuario elija «Con voz».
 */
let lecturaVozActiva = false;
const LECTURA_VOZ_KEY = 'fichaUNCP_lecturaVoz';
/** Solo true si el usuario activó voz en esta sesión (no se restaura al cargar). */
let usuarioEligioVoz = false;

// ═══════════════════════════════════════════
// VOZ (TTS + micrófono) — robusto para Chrome/Edge
// ═══════════════════════════════════════════
function contextoSeguro() {
  // El micrófono de Chrome exige https o localhost (file:// falla)
  return !!(window.isSecureContext);
}

function obtenerVoces() {
  try {
    return speechSynthesis.getVoices() || [];
  } catch (e) {
    return [];
  }
}

function obtenerVozEspañol() {
  const voces = obtenerVoces();
  if (!voces.length) return null;
  const pref = [
    v => /^es-PE/i.test(v.lang),
    v => /^es-MX/i.test(v.lang),
    v => /^es-US/i.test(v.lang),
    v => /^es-ES/i.test(v.lang),
    v => /^es/i.test(v.lang),
    v => /spanish|español/i.test(v.name || '')
  ];
  for (const fn of pref) {
    const hit = voces.find(fn);
    if (hit) return hit;
  }
  return voces[0] || null;
}

function detenerHabla() {
  hablarToken++;
  if (hablarTimerSeguridad) {
    clearTimeout(hablarTimerSeguridad);
    hablarTimerSeguridad = null;
  }
  try {
    if ('speechSynthesis' in window) {
      speechSynthesis.cancel();
      // Chrome a veces queda en paused tras cancel
      if (speechSynthesis.paused) speechSynthesis.resume();
    }
  } catch (e) { /* ignore */ }
}

/**
 * Aplica el modo de lectura (con/sin narrador) a la UI y al estado.
 * @param {boolean} activa - true = TTS lee las preguntas
 * @param {{persistir?: boolean}} [opts]
 */
function setLecturaVoz(activa, opts) {
  lecturaVozActiva = !!activa;
  if (lecturaVozActiva) usuarioEligioVoz = true;
  if (!opts || opts.persistir !== false) {
    try { localStorage.setItem(LECTURA_VOZ_KEY, lecturaVozActiva ? '1' : '0'); } catch (e) { /* ignore */ }
  }
  document.body.classList.toggle('sin-lectura-voz', !lecturaVozActiva);
  const btn = document.getElementById('btn-toggle-lectura');
  if (btn) {
    if (lecturaVozActiva) {
      btn.textContent = '🔊 Con voz';
      btn.setAttribute('aria-label', 'Desactivar lectura en voz alta');
      btn.title = 'Lectura en voz alta activada. Pulse para desactivar.';
    } else {
      btn.textContent = '🔇 Sin voz';
      btn.setAttribute('aria-label', 'Activar lectura en voz alta');
      btn.title = 'Sin narrador (por defecto). Pulse para activar lectura en voz alta.';
    }
  }
  if (!lecturaVozActiva) {
    detenerHabla();
  }
}

/** Alterna lectura en voz alta durante la entrevista (botón de la barra). */
function alternarLecturaVoz() {
  const nuevo = !lecturaVozActiva;
  setLecturaVoz(nuevo);
  if (nuevo) {
    const q = cola[indice];
    const txt = q
      ? ('Lectura en voz alta activada. ' + (q.vozPregunta || q.pregunta || ''))
      : 'Lectura en voz alta activada.';
    hablar(txt, function () {
      if (modoVozContinua && enPantallaEntrevista()) programarEscuchaAutomatica(300);
    });
  } else {
    mostrarMensaje('Narrador desactivado. Solo texto en pantalla. Puede reactivarlo con el botón Con voz.');
  }
}

/**
 * Habla texto en español (solo si lecturaVozActiva).
 * Sin narrador: muestra el texto y ejecuta alTerminar de inmediato.
 * Corrige bug de Chrome: cancel()+speak() inmediato se ignora.
 * Garantiza alTerminar con timeout de seguridad.
 */
function hablar(texto, alTerminar) {
  const msg = String(texto || '').trim();
  mostrarMensaje(msg);

  if (!msg) {
    if (alTerminar) setTimeout(alTerminar, 50);
    return;
  }

  // Modo sin voz: no narrar; el flujo de la entrevista sigue igual
  if (!lecturaVozActiva) {
    if (alTerminar) setTimeout(alTerminar, 80);
    return;
  }

  if (!('speechSynthesis' in window)) {
    if (alTerminar) setTimeout(alTerminar, 400);
    return;
  }

  detenerHabla();
  const token = hablarToken;

  // Esperar a que cancel surta efecto (bug Chrome/Windows)
  setTimeout(() => {
    if (token !== hablarToken) return;

    const u = new SpeechSynthesisUtterance(msg);
    const voz = obtenerVozEspañol();
    if (voz) {
      u.voice = voz;
      u.lang = voz.lang || 'es-ES';
    } else {
      u.lang = 'es-ES';
    }
    u.rate = Math.min(1.4, Math.max(0.6, velocidadVoz));
    u.pitch = 1;
    u.volume = 1;

    let terminado = false;
    const fin = () => {
      if (terminado || token !== hablarToken) return;
      terminado = true;
      if (hablarTimerSeguridad) {
        clearTimeout(hablarTimerSeguridad);
        hablarTimerSeguridad = null;
      }
      if (typeof alTerminar === 'function') {
        try { alTerminar(); } catch (e) { console.error(e); }
      }
    };

    u.onend = fin;
    u.onerror = (ev) => {
      console.warn('TTS error', ev && ev.error);
      fin();
    };

    // Si onend no dispara (otro bug Chrome), no bloquear la entrevista
    const ms = Math.min(90000, Math.max(2500, Math.round(msg.length * 90 / u.rate)));
    hablarTimerSeguridad = setTimeout(fin, ms);

    try {
      speechSynthesis.speak(u);
      // Chrome a veces pausa el queue en segundo plano
      if (speechSynthesis.paused) speechSynthesis.resume();
    } catch (e) {
      console.error('speak failed', e);
      fin();
    }

    // Watchdog: reanudar si se “congela” (textos largos en Chrome)
    const keepAlive = setInterval(() => {
      if (token !== hablarToken || terminado) {
        clearInterval(keepAlive);
        return;
      }
      try {
        if (speechSynthesis.speaking && speechSynthesis.paused) {
          speechSynthesis.resume();
        }
        if (!speechSynthesis.speaking && !speechSynthesis.pending) {
          clearInterval(keepAlive);
        }
      } catch (e) { clearInterval(keepAlive); }
    }, 4000);
  }, 120);
}

function mostrarMensaje(texto) {
  const el = document.getElementById('mensaje-voz');
  if (el) {
    const icono = lecturaVozActiva ? '🔊 ' : '💬 ';
    el.textContent = icono + (texto || '');
  }
  // no pisa el resumen salvo que estemos en entrevista
}

function pausarVoz() {
  detenerHabla();
  detenerMicrofono(false);
  mostrarMensaje('Lectura pausada. Puede repetir la pregunta con el botón de altavoz.');
}

function nombreVelocidadVoz() {
  if (velocidadVoz <= 0.75) return 'lenta';
  if (velocidadVoz >= 1.15) return 'rápida';
  return 'normal';
}

/** dir: -1 más lento, +1 más rápido, 0 ciclar (botón). */
function ajustarVelocidadVoz(dir) {
  if (dir < 0) {
    velocidadVoz = Math.max(0.6, Math.round((velocidadVoz - 0.2) * 100) / 100);
  } else if (dir > 0) {
    velocidadVoz = Math.min(1.4, Math.round((velocidadVoz + 0.2) * 100) / 100);
  } else {
    if (velocidadVoz < 0.95) velocidadVoz = 0.9;
    else if (velocidadVoz < 1.15) velocidadVoz = 1.2;
    else velocidadVoz = 0.7;
  }
  const nombre = nombreVelocidadVoz();
  mostrarMensaje('Velocidad: ' + nombre + ' (' + velocidadVoz.toFixed(1) + ')');
  hablar('Velocidad de voz: ' + nombre + '.', function () {
    if (modoVozContinua && enPantallaEntrevista()) programarEscuchaAutomatica(300);
  });
}

function cambiarVelocidad() {
  ajustarVelocidadVoz(0);
}

function textoAyudaVoz() {
  return (
    'El micrófono queda activo toda la entrevista: solo se pide permiso una vez. ' +
    'Después de cada pregunta, escucho solo; no hace falta pulsar Voz otra vez. ' +
    'Comandos: diga repetir u otra vez para oír la pregunta; ' +
    'anterior para volver; omitir para saltar; ' +
    'más lento o más rápido para la velocidad; ' +
    'ayuda para esta lista; apagar voz para desactivar el micrófono automático. ' +
    'En opciones, diga la respuesta y avanzaré solo. ' +
    'En texto, dicte la respuesta y diga confirmar o siguiente.'
  );
}

function enPantallaEntrevista() {
  return !!(document.getElementById('pantalla-entrevista')?.classList.contains('activa'));
}

function cancelarEscuchaProgramada() {
  if (escuchaAutoTimer) {
    clearTimeout(escuchaAutoTimer);
    escuchaAutoTimer = null;
  }
  escuchaAutoToken++;
}

/**
 * Tras leer la pregunta, si el modo continuo está activo, abre el micrófono solo.
 */
function programarEscuchaAutomatica(delayMs) {
  cancelarEscuchaProgramada();
  if (!modoVozContinua) return;
  if (!enPantallaEntrevista()) return;
  if (!soportaVoz() || !contextoSeguro()) return;

  const token = escuchaAutoToken;
  const espera = typeof delayMs === 'number' ? delayMs : 450;
  escuchaAutoTimer = setTimeout(function () {
    escuchaAutoTimer = null;
    if (token !== escuchaAutoToken) return;
    if (!modoVozContinua || !enPantallaEntrevista()) return;
    if (escuchando) return;
    // No abrir mic si aún habla el sistema
    try {
      if (speechSynthesis.speaking || speechSynthesis.pending) {
        programarEscuchaAutomatica(400);
        return;
      }
    } catch (e) { /* ignore */ }
    iniciarEscucha({ auto: true, silencioso: true });
  }, espera);
}

/**
 * Activa permiso + modo continuo (una sola vez por sesión de entrevista).
 * @returns {Promise<boolean>}
 */
async function activarModoVozContinua(opts) {
  opts = opts || {};
  if (!soportaVoz()) {
    if (!opts.silencioso) {
      const msg = 'Este navegador no soporta dictado por voz. Use Chrome o Edge.';
      mostrarMensaje(msg);
      if (!opts.sinHablar) hablar(msg);
    }
    return false;
  }
  if (!contextoSeguro()) {
    if (!opts.silencioso) {
      const msg = mensajeErrorMicro('not-secure');
      mostrarMensaje(msg);
      if (!opts.sinHablar) hablar(msg);
    }
    return false;
  }
  const perm = await asegurarPermisoMicro();
  if (!perm.ok) {
    if (!opts.silencioso) {
      const msg = mensajeErrorMicro(perm.code);
      mostrarMensaje(msg);
      if (!opts.sinHablar) hablar(msg);
    }
    return false;
  }
  const yaEstaba = modoVozContinua;
  modoVozContinua = true;
  actualizarBotonMicro();
  if (!yaEstaba && !opts.silencioso && !opts.sinHablar) {
    mostrarMensaje('✓ Voz activada. Escucharé solo después de cada pregunta. No hace falta pulsar Voz otra vez.');
  } else if (!opts.silencioso) {
    mostrarMensaje('✓ Micrófono listo (modo continuo).');
  }
  return true;
}

function desactivarModoVozContinua(anunciar) {
  modoVozContinua = false;
  cancelarEscuchaProgramada();
  detenerMicrofono(false);
  actualizarBotonMicro();
  if (anunciar !== false) {
    mostrarMensaje('Micrófono automático desactivado. Pulse 🎤 Voz para reactivar.');
    hablar('Micrófono automático desactivado. Pulse el botón Voz cuando quiera reactivarlo.');
  }
}

function aumentarFuente() {
  const a = parseFloat(getComputedStyle(document.body).fontSize);
  if (a < 34) document.body.style.fontSize = (a + 2) + 'px';
}
function disminuirFuente() {
  const a = parseFloat(getComputedStyle(document.body).fontSize);
  if (a > 14) document.body.style.fontSize = (a - 2) + 'px';
}

function leerBienvenida() {
  // Forzar TTS solo para este anuncio (instrucciones en la pantalla de inicio)
  const prev = lecturaVozActiva;
  lecturaVozActiva = true;
  hablar(
    'Bienvenida a la Ficha Social de la Unión Nacional de Ciegos del Perú. ' +
    'Al comenzar puede elegir: con voz, el sistema lee cada pregunta en voz alta; ' +
    'o sin voz, solo texto en pantalla, sin narrador. ' +
    'Se recomienda una persona de apoyo. Las respuestas se pueden escribir o dictar con el micrófono. ' +
    'Al final la ficha se guarda en Google Sheets y el PDF en Drive. ' +
    'Pulse Comenzar con voz o Comenzar sin voz.',
    function () { lecturaVozActiva = prev; }
  );
}

function soportaVoz() {
  return !!(window.SpeechRecognition || window.webkitSpeechRecognition);
}

function mensajeErrorMicro(code) {
  const map = {
    'not-allowed': 'Permiso de micrófono denegado. En el candado de la barra de dirección permita el micrófono y reintente.',
    'service-not-allowed': 'El navegador bloqueó el reconocimiento de voz. Use Chrome o Edge y abra con el archivo Abrir entrevista.bat',
    'network': 'El reconocimiento de voz necesita conexión a internet (servicio de Google/Edge).',
    'no-speech': 'No se escuchó nada. Hable más cerca del micrófono y pulse Voz otra vez.',
    'audio-capture': 'No se detectó micrófono. Revise que esté conectado y habilitado en Windows.',
    'aborted': null,
    'not-secure': 'El micrófono no funciona abriendo el HTML con doble clic (file). Use Abrir entrevista.bat para iniciar el servidor local.'
  };
  return map[code] || ('Error de micrófono: ' + code + '. Puede escribir la respuesta con el teclado.');
}

async function asegurarPermisoMicro() {
  if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
    return { ok: false, code: 'audio-capture' };
  }
  if (!contextoSeguro()) {
    return { ok: false, code: 'not-secure' };
  }
  try {
    if (microStream) {
      // verificar tracks vivas
      const vivos = microStream.getTracks().some(t => t.readyState === 'live');
      if (vivos) return { ok: true };
    }
    microStream = await navigator.mediaDevices.getUserMedia({
      audio: {
        echoCancellation: true,
        noiseSuppression: true,
        autoGainControl: true
      },
      video: false
    });
    // No usamos el stream para Web Speech API, solo para pedir permiso.
    // Mantener tracks activas ayuda a que el reconocimiento no falle por permiso.
    return { ok: true };
  } catch (e) {
    const name = (e && e.name) || '';
    if (name === 'NotAllowedError' || name === 'PermissionDeniedError') {
      return { ok: false, code: 'not-allowed' };
    }
    if (name === 'NotFoundError' || name === 'DevicesNotFoundError') {
      return { ok: false, code: 'audio-capture' };
    }
    return { ok: false, code: 'audio-capture' };
  }
}

function limpiarTimerSilencio() {
  if (silencioTimer) {
    clearTimeout(silencioTimer);
    silencioTimer = null;
  }
}

function programarFinPorSilencio() {
  limpiarTimerSilencio();
  // Tras ~1,3 s sin más palabras finales, procesar y avanzar (no hace falta pulsar Detener)
  silencioTimer = setTimeout(function () {
    silencioTimer = null;
    const texto = (transcriptFinal || transcriptParcial || '').trim();
    if (!texto) return;
    detenerMicrofono(false);
    procesarDictadoVoz(texto);
  }, 1300);
}

function crearReconocimiento() {
  const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
  if (!SR) return null;
  const r = new SR();
  // es-ES es el más estable en Chrome; es-PE a menudo no está disponible
  r.lang = 'es-ES';
  r.continuous = true;
  r.interimResults = true;
  r.maxAlternatives = 3;

  r.onstart = () => {
    escuchando = true;
    actualizarBotonMicro();
    mostrarMensaje('Escuchando… hable ahora. Al callar, se procesa solo. También puede pulsar Detener.');
  };

  r.onresult = (ev) => {
    let interim = '';
    let huboFinal = false;
    for (let i = ev.resultIndex; i < ev.results.length; i++) {
      const piece = ev.results[i][0].transcript;
      if (ev.results[i].isFinal) {
        transcriptFinal = (transcriptFinal + ' ' + piece).trim();
        huboFinal = true;
      } else {
        interim += piece;
      }
    }
    const vista = (transcriptFinal + (interim ? ' ' + interim : '')).trim();
    transcriptParcial = vista;
    mostrarMensaje('🎤 ' + (vista || 'Escuchando…'));

    if (huboFinal && transcriptFinal) {
      // Comando global completo → ejecutar al instante
      const cmd = detectarComandoVoz(transcriptFinal);
      if (cmd && cmd.inmediato) {
        const texto = transcriptFinal;
        detenerMicrofono(false);
        procesarDictadoVoz(texto);
        return;
      }
      programarFinPorSilencio();
    }
  };

  r.onerror = (ev) => {
    const code = (ev && ev.error) || 'unknown';
    console.warn('STT error', code);
    if (code === 'aborted' || code === 'no-speech') {
      if (code === 'no-speech' && !transcriptFinal && !transcriptParcial) {
        mostrarMensaje(mensajeErrorMicro('no-speech'));
      }
      return;
    }
    limpiarTimerSilencio();
    escuchando = false;
    actualizarBotonMicro();
    const msg = mensajeErrorMicro(code);
    if (msg) {
      mostrarMensaje(msg);
      if (code !== 'network' && code !== 'not-allowed' && code !== 'service-not-allowed') {
        setTimeout(() => hablar(msg), 200);
      }
    }
  };

  r.onend = () => {
    limpiarTimerSilencio();
    const estaba = escuchando;
    escuchando = false;
    actualizarBotonMicro();
    const texto = (transcriptFinal || transcriptParcial || '').trim();
    transcriptFinal = '';
    transcriptParcial = '';
    if (estaba && texto) {
      procesarDictadoVoz(texto);
    } else if (estaba && !texto) {
      mostrarMensaje(
        modoVozContinua
          ? 'No se captó audio. Vuelvo a escuchar… diga la respuesta o un comando.'
          : 'No se captó audio. Pulse 🎤 Activar voz e intente de nuevo, o escriba la respuesta.'
      );
      if (modoVozContinua && enPantallaEntrevista()) {
        programarEscuchaAutomatica(900);
      }
    }
  };

  return r;
}

/**
 * Detiene el micrófono.
 * @param {boolean} aplicar  Si true, devuelve el texto dictado para aplicarlo al campo.
 * @returns {string} texto pendiente de dictado (si había)
 */
function detenerMicrofono(aplicar = true) {
  limpiarTimerSilencio();
  const textoPendiente = (transcriptFinal || transcriptParcial || '').trim();
  transcriptFinal = '';
  transcriptParcial = '';

  if (reconocimiento) {
    try {
      // abort evita que onend vuelva a aplicar el mismo texto
      reconocimiento.onresult = null;
      reconocimiento.onerror = null;
      reconocimiento.onend = null;
      reconocimiento.abort();
    } catch (e) { /* ignore */ }
    reconocimiento = null;
  }
  escuchando = false;
  actualizarBotonMicro();
  return aplicar ? textoPendiente : '';
}

/**
 * Abre el reconocimiento de voz.
 * @param {{ auto?: boolean, silencioso?: boolean }} opts
 *   auto: llamado tras cada pregunta (no corta el TTS si ya terminó)
 *   silencioso: menos mensajes
 */
async function iniciarEscucha(opts) {
  opts = opts || {};
  if (!soportaVoz()) {
    if (!opts.silencioso) {
      const msg = 'Este navegador no soporta dictado por voz. Use Google Chrome o Microsoft Edge, o escriba la respuesta.';
      mostrarMensaje(msg);
      hablar(msg);
    }
    return false;
  }
  if (escuchando) return true;

  if (!opts.auto) {
    // Clic manual: cortar TTS para no grabarlo
    detenerHabla();
  }
  if (vozAutoAvanceTimer) {
    clearTimeout(vozAutoAvanceTimer);
    vozAutoAvanceTimer = null;
  }

  const ok = await activarModoVozContinua({
    silencioso: !!opts.auto || !!opts.silencioso,
    sinHablar: true
  });
  if (!ok) return false;

  try {
    if (reconocimiento) {
      try { reconocimiento.abort(); } catch (e) { /* ignore */ }
    }
  } catch (e) { /* ignore */ }

  transcriptFinal = '';
  transcriptParcial = '';
  reconocimiento = crearReconocimiento();
  if (!reconocimiento) {
    if (!opts.silencioso) hablar('No se pudo crear el reconocimiento de voz.');
    return false;
  }

  // Pausa breve: liberar altavoz antes de abrir mic (más larga si viene del auto)
  await new Promise(r => setTimeout(r, opts.auto ? 280 : 350));

  // Si mientras esperábamos cambió la pregunta o se apagó el modo, no abrir
  if (opts.auto && (!modoVozContinua || !enPantallaEntrevista())) return false;
  try {
    if (speechSynthesis.speaking || speechSynthesis.pending) {
      // El sistema aún habla: reintentar luego
      if (opts.auto) programarEscuchaAutomatica(500);
      return false;
    }
  } catch (e) { /* ignore */ }

  try {
    reconocimiento.start();
    escuchando = true;
    actualizarBotonMicro();
    if (opts.auto) {
      mostrarMensaje('🎤 Escuchando… diga la respuesta o un comando. No hace falta pulsar Voz.');
    } else {
      mostrarMensaje('🎤 Escuchando… diga la respuesta o un comando (repetir, anterior, omitir, ayuda…).');
    }
    return true;
  } catch (e) {
    console.error(e);
    escuchando = false;
    actualizarBotonMicro();
    if (!opts.silencioso) {
      const msg = 'No se pudo iniciar el micrófono. Cierre otras apps que lo usen, permita el permiso y reintente.';
      mostrarMensaje(msg);
      hablar(msg);
    } else {
      // Reintento suave en modo auto
      programarEscuchaAutomatica(1200);
    }
    return false;
  }
}

/**
 * Botón 🎤:
 * - Si está escuchando → detiene y procesa lo dictado (el modo continuo sigue activo).
 * - Si no → activa permiso (1 vez) + modo continuo + empieza a escuchar.
 * - Con modo continuo ya activo, no hace falta usarlo: escucha solo tras cada pregunta.
 */
async function toggleMicrofono() {
  if (!soportaVoz()) {
    const msg = 'Este navegador no soporta dictado por voz. Use Google Chrome o Microsoft Edge, o escriba la respuesta.';
    mostrarMensaje(msg);
    hablar(msg);
    return;
  }

  // Si ya escucha → detener y procesar (modo continuo se mantiene)
  if (escuchando) {
    const texto = detenerMicrofono(true);
    if (texto) procesarDictadoVoz(texto);
    else {
      mostrarMensaje('No se captó audio. Espere la pregunta o diga la respuesta cuando escuche «Escuchando».');
      if (modoVozContinua) programarEscuchaAutomatica(600);
    }
    return;
  }

  const primeraVez = !modoVozContinua;
  const ok = await iniciarEscucha({ auto: false, silencioso: false });
  if (ok && primeraVez) {
    hablar(
      'Voz activada. A partir de ahora escucharé solo después de cada pregunta. No hace falta volver a pulsar el botón de voz. Diga apagar voz si quiere desactivarlo.'
    );
  }
}

function actualizarBotonMicro() {
  const btn = document.getElementById('btn-microfono');
  if (!btn) return;
  if (escuchando) {
    btn.classList.add('escuchando');
    btn.classList.remove('voz-continua');
    btn.textContent = '⏹ Detener';
    btn.setAttribute('aria-pressed', 'true');
    btn.title = 'Detener escucha de esta respuesta (el modo voz sigue activo)';
  } else if (modoVozContinua) {
    btn.classList.remove('escuchando');
    btn.classList.add('voz-continua');
    btn.textContent = '🎤 Voz ON';
    btn.setAttribute('aria-pressed', 'false');
    btn.title = 'Voz continua activa: escucha sola tras cada pregunta. Pulse para escuchar ya.';
  } else {
    btn.classList.remove('escuchando');
    btn.classList.remove('voz-continua');
    btn.textContent = '🎤 Activar voz';
    btn.setAttribute('aria-pressed', 'false');
    btn.title = 'Pulse una vez para activar el micrófono. Luego escucha sola en cada pregunta.';
  }
}

function actualizarAvisoVozUI() {
  const zona = document.getElementById('aviso-voz');
  if (!zona) return;
  const partes = [];
  if (!('speechSynthesis' in window)) {
    partes.push('Este navegador no lee en voz alta.');
  }
  if (!soportaVoz()) {
    partes.push('Dictado no disponible. Use Chrome o Edge.');
  } else if (!contextoSeguro()) {
    partes.push('⚠️ Micrófono bloqueado en modo archivo. Ejecute «Abrir entrevista.bat» (servidor local) para usar el dictado.');
  } else if (modoVozContinua) {
    partes.push('✓ Voz continua activa: escucha sola después de cada pregunta. Comandos: repetir, anterior, omitir, confirmar, apagar voz, ayuda.');
  } else {
    partes.push('Pulse «Comenzar» o «🎤 Activar voz» una sola vez (permiso del micrófono). Luego escucha sola en cada pregunta.');
  }
  zona.textContent = partes.join(' ');
  zona.hidden = false;
}

function normalizar(s) {
  return (s || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
}

/** ¿La pregunta actual es sí/no (para no confundir "sí" con confirmar)? */
function preguntaEsSiNo(q) {
  if (!q || !q.opciones) return false;
  const vals = q.opciones.map(o => o.value);
  return vals.includes('si') && vals.includes('no');
}

/**
 * Detecta comandos globales de voz.
 * @returns {{ cmd: string, inmediato: boolean, resto?: string } | null}
 */
function detectarComandoVoz(texto) {
  const n = normalizar(texto);
  if (!n) return null;

  // Exactos (ejecutar al instante)
  const exactos = [
    { cmd: 'repetir', re: /^(repetir|repite|repite por favor|otra vez|de nuevo|repetir pregunta|vuelve a leer)$/ },
    { cmd: 'anterior', re: /^(anterior|atras|atras|volver|pregunta anterior|regresar)$/ },
    { cmd: 'omitir', re: /^(omitir|saltar|pasar|no se|no se la respuesta|prefiero no responder|sin respuesta|no aplica)$/ },
    { cmd: 'mas_lento', re: /^(mas lento|mas despacio|hablar mas lento|mas despacio por favor|velocidad lenta)$/ },
    { cmd: 'mas_rapido', re: /^(mas rapido|mas deprisa|hablar mas rapido|velocidad rapida|mas rapido por favor)$/ },
    { cmd: 'ayuda', re: /^(ayuda|comandos|ayuda por favor|que puedo decir|ayuda de voz)$/ },
    { cmd: 'confirmar', re: /^(confirmar|siguiente|continuar|listo|ok|okay|okei|de acuerdo|correcto|adelante|siguiente pregunta)$/ },
    { cmd: 'apagar_voz', re: /^(apagar voz|apagar microfono|apagar micrófono|desactivar voz|desactivar microfono|desactivar micrófono|silencio microfono|no escuchar)$/ }
  ];
  for (const e of exactos) {
    if (e.re.test(n)) return { cmd: e.cmd, inmediato: true };
  }

  // "sí" / "si" solo como confirmar si NO es pregunta sí/no
  if (/^(si|sí|yes)$/.test(n) || n === 'si' || n === 'yes') {
    const q = cola[indice];
    if (!preguntaEsSiNo(q)) {
      return { cmd: 'confirmar', inmediato: true };
    }
  }

  return null;
}

/** Quita "confirmar/siguiente/…" al final del dictado de una respuesta. */
function separarConfirmacionDelTexto(texto) {
  const raw = String(texto || '').trim();
  const re = /\s+(confirmar|siguiente|continuar|listo|ok|okay|de acuerdo|correcto|adelante)\s*$/i;
  const m = raw.match(re);
  if (m) {
    return { texto: raw.slice(0, m.index).trim(), confirmar: true };
  }
  return { texto: raw, confirmar: false };
}

function hayValorRespuestaActual() {
  const q = cola[indice];
  if (!q) return false;
  const v = leerRespuestaActual();
  if (Array.isArray(v)) return v.length > 0;
  return v !== '' && v !== null && v !== undefined;
}

function ejecutarComandoVoz(cmd) {
  const enEntrevista = document.getElementById('pantalla-entrevista')?.classList.contains('activa');

  switch (cmd) {
    case 'repetir':
      if (enEntrevista) repetirPregunta();
      else hablar('No hay pregunta activa. Comience la entrevista.');
      break;
    case 'anterior':
      if (enEntrevista) {
        mostrarMensaje('Comando: anterior');
        preguntaAnterior();
      } else {
        hablar('No hay pregunta activa.');
      }
      break;
    case 'omitir':
      if (enEntrevista) {
        mostrarMensaje('Comando: omitir');
        saltarPregunta();
      } else {
        hablar('No hay pregunta activa.');
      }
      break;
    case 'mas_lento':
      ajustarVelocidadVoz(-1);
      break;
    case 'mas_rapido':
      ajustarVelocidadVoz(1);
      break;
    case 'ayuda':
      hablar(textoAyudaVoz(), function () {
        if (modoVozContinua && enEntrevista) programarEscuchaAutomatica(300);
      });
      mostrarMensaje(textoAyudaVoz());
      break;
    case 'confirmar':
      if (!enEntrevista) {
        hablar('No hay pregunta activa.');
        break;
      }
      if (hayValorRespuestaActual()) {
        mostrarMensaje('Comando: confirmar → avanzando');
        confirmarYSiguiente();
      } else {
        hablar('Aún no hay respuesta. Díctela primero o diga omitir.', function () {
          if (modoVozContinua) programarEscuchaAutomatica(300);
        });
        mostrarMensaje('⚠ Sin respuesta para confirmar. Dicte o diga omitir.');
      }
      break;
    case 'apagar_voz':
      desactivarModoVozContinua(true);
      break;
    default:
      break;
  }
}

/**
 * Punto único: comandos globales + respuesta + auto-avance.
 */
function procesarDictadoVoz(texto) {
  const raw = String(texto || '').trim();
  if (!raw) return;
  if (procesandoDictado) return;
  procesandoDictado = true;

  try {
    mostrarMensaje('Escuchado: ' + raw);

    // 1) Comando global puro
    const cmd = detectarComandoVoz(raw);
    if (cmd && cmd.inmediato) {
      ejecutarComandoVoz(cmd.cmd);
      return;
    }

    const enEntrevista = document.getElementById('pantalla-entrevista')?.classList.contains('activa');
    if (!enEntrevista) {
      hablar('Comience la entrevista para dictar respuestas. Diga ayuda para oír los comandos.');
      return;
    }

    // 2) Respuesta + posible "confirmar" al final
    const sep = separarConfirmacionDelTexto(raw);
    const cuerpo = sep.texto || raw;
    const forzarConfirmar = sep.confirmar;

    const result = aplicarRespuestaVoz(cuerpo, { anunciar: false });
    if (!result || !result.ok) {
      // Si no hubo match pero el usuario dijo solo confirmar y ya hay valor
      if (forzarConfirmar && hayValorRespuestaActual()) {
        confirmarYSiguiente();
        return;
      }
      // Reabrir mic para reintentar (modo continuo), tras el mensaje de error
      if (modoVozContinua && enPantallaEntrevista()) {
        programarEscuchaAutomatica(2200);
      }
      return;
    }

    // Opciones / número / checkbox claros → avanzar solo
    // Texto libre → avanzar si dijo confirmar, o también si es texto corto ya capturado y forzarConfirmar
    const auto = result.autoAvanzar || forzarConfirmar;
    const etiqueta = result.label || cuerpo;

    if (auto) {
      mostrarMensaje('✓ ' + etiqueta + ' → siguiente');
      if (vozAutoAvanceTimer) clearTimeout(vozAutoAvanceTimer);
      const idxAntes = indice;
      let yaAvanzo = false;
      const avanzar = function () {
        if (yaAvanzo) return;
        yaAvanzo = true;
        if (vozAutoAvanceTimer) {
          clearTimeout(vozAutoAvanceTimer);
          vozAutoAvanceTimer = null;
        }
        // Solo avanzar si seguimos en la misma pregunta
        if (indice === idxAntes && !_confirmando) {
          confirmarYSiguiente();
        }
      };
      // Breve anuncio y avanzar (un solo paso)
      hablar('Registrado: ' + etiqueta, avanzar);
      // Si TTS no dispara onend, avanzar igual
      vozAutoAvanceTimer = setTimeout(avanzar, 2800);
    } else {
      mostrarMensaje('✓ ' + etiqueta + '. Diga confirmar o siguiente para continuar.');
      hablar('Registrado: ' + etiqueta + '. Diga confirmar o siguiente para continuar.', function () {
        if (modoVozContinua && enPantallaEntrevista()) programarEscuchaAutomatica(300);
      });
    }
  } finally {
    // Liberar tras un tick para permitir el flujo de confirmar
    setTimeout(function () { procesandoDictado = false; }, 300);
  }
}

/** Escribe valor en control visible y en respuestas[] (persistencia inmediata). */
function escribirValorEnCampo(q, valor, anunciar) {
  if (anunciar === undefined) anunciar = true;
  if (!q || valor === undefined || valor === null) return false;

  if (q.tipo === 'radio') {
    document.querySelectorAll('.opcion').forEach(o => o.classList.remove('seleccionada'));
    const radio = document.querySelector('input[name="resp"][value="' + CSS.escape(String(valor)) + '"]');
    if (radio) {
      radio.checked = true;
      radio.closest('.opcion')?.classList.add('seleccionada');
    }
    respuestas[q.id] = valor;
    if (anunciar) {
      const lab = (q.opciones || []).find(o => o.value === valor)?.label || valor;
      mostrarMensaje('✓ ' + lab + '. Diga confirmar o espere el avance automático.');
    }
    return true;
  }

  if (q.tipo === 'select') {
    const sel = document.getElementById('campo-resp');
    if (sel) sel.value = valor;
    respuestas[q.id] = valor;
    if (anunciar) {
      const lab = (q.opciones || []).find(o => o.value === valor)?.label || valor;
      mostrarMensaje('✓ ' + lab);
    }
    return true;
  }

  if (q.tipo === 'checkbox') {
    const arr = Array.isArray(valor)
      ? valor
      : String(valor).split(',').map(s => s.trim()).filter(Boolean);
    document.querySelectorAll('input[name="resp"]').forEach(c => { c.checked = false; });
    arr.forEach(v => {
      const c = document.querySelector('input[name="resp"][value="' + CSS.escape(v) + '"]');
      if (c) c.checked = true;
    });
    respuestas[q.id] = arr;
    if (anunciar) mostrarMensaje('✓ Opciones marcadas.');
    return true;
  }

  const el = document.getElementById('campo-resp');
  if (el) el.value = valor;
  respuestas[q.id] = valor;
  if (anunciar) mostrarMensaje('✓ Guardado: ' + valor);
  return true;
}

/**
 * Aplica dictado a la pregunta actual.
 * @returns {{ ok: boolean, autoAvanzar?: boolean, label?: string }}
 */
function aplicarRespuestaVoz(texto, opts) {
  opts = opts || {};
  const anunciar = opts.anunciar !== false;
  const q = cola[indice];
  if (!q) return { ok: false };
  const raw = String(texto || '').trim();
  if (!raw) return { ok: false };

  if (q.tipo === 'radio' || q.tipo === 'select') {
    const n = normalizar(raw);
    let match = null;
    for (const op of q.opciones) {
      const ln = normalizar(op.label);
      const vn = normalizar(op.value);
      if (n === ln || n === vn || n.includes(ln) || ln.includes(n) || n.includes(vn)) {
        match = op;
        break;
      }
    }
    if (!match) {
      const mapa = {
        si: ['si', 'sí', 'yes'], no: ['no'],
        M: ['masculino', 'hombre', 'varon', 'varón', 'eme'],
        F: ['femenino', 'mujer', 'efe'],
        soltero: ['soltero', 'soltera'], casado: ['casado', 'casada'],
        viudo: ['viudo', 'viuda'], conviviente: ['conviviente', 'convive'],
        divorciado: ['divorciado', 'divorciada'], separado: ['separado', 'separada'],
        nuclear: ['nuclear'], extensa: ['extensa'], unipersonal: ['unipersonal', 'solo', 'sola'],
        monoparental: ['monoparental'],
        ceguera: ['ceguera', 'ciego', 'ciega', 'total'], 'baja-vision': ['baja vision', 'baja visión'],
        congenita: ['congenita', 'congénita', 'nacimiento'], adquirida: ['adquirida'],
        sis: ['sis'], essalud: ['essalud', 'es salud'], privado: ['privado'],
        independiente: ['independiente'], parcial: ['parcial', 'apoyo parcial'],
        permanente: ['permanente', 'apoyo permanente'],
        dependiente: ['dependiente'], pension: ['pension', 'pensión', 'jubilacion', 'jubilación'],
        programa: ['programa'], 'apoyo-familiar': ['apoyo familiar', 'familia'],
        propia: ['propia'], alquilada: ['alquilada', 'alquilo'], cedida: ['cedida'],
        casa: ['casa'], departamento: ['departamento'], cuarto: ['cuarto'],
        quinta: ['quinta', 'callejon', 'callejón'], multifamiliar: ['multifamiliar'],
        propio: ['propio', 'dormitorio propio'], 'comparte-dorm': ['comparte dormitorio'],
        'comparte-hab': ['comparte habitacion', 'comparte habitación'],
        pobre: ['pobre'], 'pobre-extremo': ['extremo', 'pobre extremo'],
        'no-pobre': ['no pobre'], 'no-tiene': ['no tiene'], 'no-sabe': ['no sabe'],
        'sin-ingresos': ['sin ingresos', 'cero'], 'menos-500': ['menos de 500', 'menos de quinientos'],
        '500-900': ['500 a 900', 'entre 500'], '900-1500': ['900 a 1500', 'entre 900'],
        'mas-1500': ['mas de 1500', 'más de 1500'],
        buena: ['buena', 'bueno'], regular: ['regular'], mala: ['mala', 'malo'],
        'sin-estudios': ['sin estudios', 'ninguno'], primaria: ['primaria incompleta', 'primaria'],
        'primaria-completa': ['primaria completa'], secundaria: ['secundaria incompleta', 'secundaria'],
        'secundaria-completa': ['secundaria completa'], tecnico: ['tecnico', 'técnico'],
        universitario: ['universitario', 'universidad']
      };
      for (const op of q.opciones) {
        const keys = mapa[op.value] || [];
        if (keys.some(k => n.includes(normalizar(k)))) { match = op; break; }
      }
    }
    if (match) {
      escribirValorEnCampo(q, match.value, false);
      // Opciones reconocidas → un solo paso (auto-avance)
      return { ok: true, autoAvanzar: true, label: match.label };
    }
    mostrarMensaje('No reconocí la opción: ' + raw);
    hablar('No reconocí la opción. Escuché: ' + raw + '. Intente de nuevo o diga ayuda.');
    return { ok: false };
  }

  if (q.tipo === 'checkbox') {
    const n = normalizar(raw);
    const elegidos = [];
    q.opciones.forEach(op => {
      if (n.includes(normalizar(op.label)) || n.includes(normalizar(op.value))) {
        elegidos.push(op.value);
      }
    });
    // Sinónimos servicios
    if (n.includes('agua')) elegidos.push('agua');
    if (n.includes('desague') || n.includes('desagüe') || n.includes('alcantarillado')) elegidos.push('desague');
    if (n.includes('luz') || n.includes('electric') || n.includes('energia') || n.includes('energía')) elegidos.push('luz');
    if (n.includes('internet') || n.includes('wifi')) elegidos.push('internet');
    const validIds = new Set((q.opciones || []).map(o => o.value));
    const uniq = [...new Set(elegidos)].filter(v => validIds.has(v));
    if (uniq.length) {
      escribirValorEnCampo(q, uniq, false);
      const lab = uniq.map(v => q.opciones.find(o => o.value === v)?.label || v).join(', ');
      return { ok: true, autoAvanzar: true, label: lab };
    }
    hablar('No reconocí las opciones. Diga por ejemplo agua, desagüe, luz e internet.');
    return { ok: false };
  }

  if (q.tipo === 'numero') {
    const digitos = raw.replace(/[^\d]/g, '');
    let num = digitos;
    if (!num) {
      const map = {
        cero: 0, uno: 1, una: 1, dos: 2, tres: 3, cuatro: 4, cinco: 5,
        seis: 6, siete: 7, ocho: 8, nueve: 9, diez: 10,
        once: 11, doce: 12, trece: 13, catorce: 14, quince: 15,
        veinte: 20, treinta: 30, cuarenta: 40, cincuenta: 50
      };
      for (const w of Object.keys(map)) {
        if (normalizar(raw).includes(w)) { num = String(map[w]); break; }
      }
    }
    if (num !== '' && num != null) {
      // Si el número viene con muchos dígitos de más (ruido), tomar el primero razonable
      if (String(num).length > 3 && q.max != null && q.max <= 120) {
        const m = String(num).match(/^\d{1,3}/);
        if (m) num = m[0];
      }
      escribirValorEnCampo(q, num, false);
      return { ok: true, autoAvanzar: true, label: String(num) };
    }
    hablar('No capté un número. Intente de nuevo o escríbalo.');
    return { ok: false };
  }

  // Texto, fecha, tel, textarea: guardar; auto-avance solo si el usuario dijo confirmar (lo maneja el caller)
  escribirValorEnCampo(q, raw, false);
  return { ok: true, autoAvanzar: false, label: raw };
}

// ═══════════════════════════════════════════
// FLUJO
// ═══════════════════════════════════════════
/**
 * Cola fija alineada a la ficha:
 * … → cantidad de convivientes → datos de cada miembro (hasta 7) → tipo de familia → …
 * Las filas de miembros fuera de la cantidad se omiten en la navegación (esPreguntaActiva).
 */
function construirColaBase() {
  const out = [];
  for (const p of PREGUNTAS) {
    out.push({ ...p });
    if (p.id === 'num_familia') {
      for (let n = 1; n <= MAX_MIEMBROS_FAMILIA; n++) {
        for (const q of plantillaMiembro(n)) {
          out.push({ ...q });
        }
      }
    }
  }
  return out;
}

/** Compat: la cola ya incluye todos los miembros; solo actualiza el contador. */
function insertarMiembrosEnCola(cantidad) {
  miembrosInsertados = Math.max(0, Math.min(MAX_MIEMBROS_FAMILIA, parseInt(cantidad, 10) || 0));
}

/**
 * Inicia la entrevista.
 * @param {boolean} [conLecturaVoz=false] - true solo si el usuario elige «Con voz». Por defecto NO suena.
 */
async function iniciarEntrevista(conLecturaVoz) {
  // Por defecto siempre sin narrador (false). Solo true si el botón «Con voz» lo pide.
  conLecturaVoz = conLecturaVoz === true;
  setLecturaVoz(conLecturaVoz);

  cancelarEscuchaProgramada();
  detenerMicrofono(false);
  detenerHabla();
  respuestas = {};
  cola = construirColaBase();
  indice = 0;
  miembrosInsertados = 0;
  respuestas.fecha_entrevista = new Date().toISOString().slice(0, 10);
  document.getElementById('pantalla-inicio').classList.remove('activa');
  document.getElementById('pantalla-resumen').classList.remove('activa');
  document.getElementById('pantalla-entrevista').classList.add('activa');
  const btn = document.getElementById('btn-microfono');
  if (btn) {
    if (!soportaVoz()) {
      btn.disabled = true;
      btn.title = 'Dictado no disponible. Use Chrome o Edge.';
      btn.textContent = '🎤 Voz no disponible';
    } else if (!contextoSeguro()) {
      btn.disabled = false;
      btn.title = 'Abra con Abrir entrevista.bat para usar el micrófono (localhost).';
    } else {
      btn.disabled = false;
    }
  }

  // Solo en modo con voz: pedir micrófono al comenzar (gesto de usuario).
  // En modo sin voz no se activa el micrófono solo; el usuario puede pulsar 🎤 si lo desea.
  let vozLista = modoVozContinua;
  if (lecturaVozActiva && !vozLista && soportaVoz() && contextoSeguro()) {
    vozLista = await activarModoVozContinua({ silencioso: true, sinHablar: true });
  }
  actualizarBotonMicro();
  actualizarAvisoVozUI();

  let introVoz;
  if (!lecturaVozActiva) {
    introVoz =
      'Comenzamos sin narrador. Las preguntas se muestran en pantalla. ' +
      'Responda con el teclado o tocando las opciones. ' +
      'Si desea dictar, pulse Activar voz. ' +
      'Puede activar la lectura en voz alta con el botón de la barra superior. ' +
      'Primera pregunta.';
  } else if (vozLista) {
    introVoz =
      'Comenzamos la entrevista con lectura en voz alta. El micrófono ya está activo: después de cada pregunta escucharé solo, ' +
      'sin que pulse el botón de voz. En opciones diga la respuesta y avanzaré solo. ' +
      'Comandos: repetir, anterior, omitir, más lento, más rápido, confirmar, ayuda, o apagar voz. ' +
      'Primera pregunta.';
  } else {
    introVoz =
      'Comenzamos la entrevista con lectura en voz alta. ' +
      'Si quiere responder por voz, pulse una sola vez el botón Activar voz para dar permiso al micrófono; ' +
      'después escucharé sola en cada pregunta. ' +
      'Comandos: repetir, anterior, omitir, confirmar, ayuda. ' +
      'Primera pregunta.';
  }

  hablar(introVoz, () => mostrarPreguntaActual());
}

function contarPreguntasActivas() {
  let total = 0;
  for (let i = 0; i < cola.length; i++) {
    if (esPreguntaActiva(cola[i])) total++;
  }
  return total;
}

function numeroPreguntaActiva(idx) {
  let n = 0;
  for (let i = 0; i <= idx && i < cola.length; i++) {
    if (esPreguntaActiva(cola[i])) n++;
  }
  return n;
}

function mostrarPreguntaActual() {
  if (indice >= cola.length) {
    irAResumen();
    return;
  }
  // Si el índice cayó en un miembro fuera de cantidad, saltar al activo
  if (!esPreguntaActiva(cola[indice])) {
    const adj = buscarIndiceActivo(indice, +1);
    if (adj >= cola.length) {
      irAResumen();
      return;
    }
    indice = adj;
  }

  cancelarEscuchaProgramada();
  detenerMicrofono(false);
  const q = cola[indice];
  const totalActivas = contarPreguntasActivas();
  const numActiva = numeroPreguntaActiva(indice);

  document.getElementById('seccion-etiqueta').textContent = SECCIONES[q.seccion] || '';
  document.getElementById('pregunta-numero').textContent =
    'Pregunta ' + numActiva + ' de ' + totalActivas;
  document.getElementById('pregunta-texto').textContent = q.pregunta;
  document.getElementById('pregunta-ayuda').textContent = q.ayuda || '';

  const pct = Math.round(((numActiva - 1) / Math.max(totalActivas, 1)) * 100);
  document.getElementById('progreso-texto').textContent =
    'Pregunta ' + numActiva + ' de ' + totalActivas + ' — ' + (SECCIONES[q.seccion] || '');
  document.getElementById('progreso-fill').style.width = pct + '%';
  document.getElementById('progreso-bar').setAttribute('aria-valuenow', pct);
  document.getElementById('btn-anterior').style.visibility = indice === 0 ? 'hidden' : 'visible';

  renderCampo(q);
  const valorGuardado = respuestas[q.id];
  if (valorGuardado !== undefined && valorGuardado !== null && valorGuardado !== '') {
    rellenarCampo(q, valorGuardado);
  }

  hablar(q.vozPregunta || q.pregunta, () => {
    const campo = document.getElementById('campo-resp') || document.querySelector('input[name="resp"]');
    if (campo) {
      try { campo.focus({ preventScroll: false }); } catch (e) { campo.focus(); }
    }
    // Modo continuo: abrir micrófono solo (sin pulsar Voz)
    programarEscuchaAutomatica(400);
  });
}

function renderCampo(q) {
  const zona = document.getElementById('zona-respuesta');
  let html = '';

  if (q.tipo === 'radio') {
    html = '<div class="opciones" role="radiogroup" aria-label="' + esc(q.pregunta) + '">';
    q.opciones.forEach((op, i) => {
      html += `<label class="opcion">
        <input type="radio" name="resp" value="${esc(op.value)}" id="op-${i}"
               onchange="marcarOpcion(this)">
        <span>${esc(op.label)}</span>
      </label>`;
    });
    html += '</div>';
  } else if (q.tipo === 'checkbox') {
    html = '<div class="opciones checks-grid">';
    q.opciones.forEach((op, i) => {
      html += `<label class="opcion">
        <input type="checkbox" name="resp" value="${esc(op.value)}" id="ck-${i}">
        <span>${esc(op.label)}</span>
      </label>`;
    });
    html += '</div>';
  } else if (q.tipo === 'select') {
    html = '<label for="campo-resp">Seleccione una opción</label><select id="campo-resp">';
    html += '<option value="">— Elija —</option>';
    q.opciones.forEach(op => {
      html += `<option value="${esc(op.value)}">${esc(op.label)}</option>`;
    });
    html += '</select>';
  } else if (q.tipo === 'textarea') {
    html = '<label for="campo-resp">Escriba o dicte la respuesta</label>' +
      '<textarea id="campo-resp" rows="4"></textarea>';
  } else if (q.tipo === 'numero') {
    html = '<label for="campo-resp">Número</label>' +
      `<input type="number" id="campo-resp" min="${q.min ?? 0}" max="${q.max ?? 999}" inputmode="numeric">`;
  } else if (q.tipo === 'fecha') {
    html = '<label for="campo-resp">Fecha</label>' +
      '<input type="date" id="campo-resp">';
  } else if (q.tipo === 'tel') {
    html = '<label for="campo-resp">Teléfono</label>' +
      '<input type="tel" id="campo-resp" inputmode="tel">';
  } else {
    html = '<label for="campo-resp">Respuesta</label>' +
      `<input type="text" id="campo-resp" ${q.maxlength ? 'maxlength="' + q.maxlength + '"' : ''} ${q.inputmode ? 'inputmode="' + q.inputmode + '"' : ''}>`;
  }
  zona.innerHTML = html;

  // Guardado en vivo al escribir / elegir (evita perder datos si no se pulsa bien Confirmar)
  const campo = document.getElementById('campo-resp');
  if (campo) {
    const persistir = () => {
      const qq = cola[indice];
      if (!qq) return;
      const v = String(campo.value || '').trim();
      if (v !== '') {
        respuestas[qq.id] = v;
        try { guardarBorrador(true); } catch (e) { /* ignore */ }
      }
    };
    campo.addEventListener('change', persistir);
    campo.addEventListener('blur', persistir);
    campo.addEventListener('keydown', (ev) => {
      if (ev.key === 'Enter' && campo.tagName !== 'TEXTAREA') {
        ev.preventDefault();
        persistir();
        confirmarYSiguiente();
      }
    });
  }
  document.querySelectorAll('input[name="resp"][type="checkbox"]').forEach(ck => {
    ck.addEventListener('change', () => {
      const qq = cola[indice];
      if (!qq) return;
      respuestas[qq.id] = Array.from(document.querySelectorAll('input[name="resp"]:checked')).map(c => c.value);
      try { guardarBorrador(true); } catch (e) { /* ignore */ }
    });
  });
}


function marcarOpcion(el) {
  document.querySelectorAll('.opcion').forEach(o => o.classList.remove('seleccionada'));
  el.closest('.opcion')?.classList.add('seleccionada');
  el.checked = true;
  const q = cola[indice];
  if (q) {
    respuestas[q.id] = el.value;
    try { guardarBorrador(true); } catch (e) { /* ignore */ }
  }
  const label = el.closest('label')?.querySelector('span')?.textContent || el.value;
  mostrarMensaje('✓ Seleccionado: ' + label + '. Pulse Confirmar para continuar.');
  hablar('Seleccionado: ' + label);
}

function rellenarCampo(q, valor) {
  if (q.tipo === 'radio') {
    const r = document.querySelector(`input[name="resp"][value="${CSS.escape(String(valor))}"]`);
    if (r) { r.checked = true; r.closest('.opcion')?.classList.add('seleccionada'); }
  } else if (q.tipo === 'checkbox') {
    const arr = Array.isArray(valor) ? valor : String(valor).split(',');
    arr.forEach(v => {
      const c = document.querySelector(`input[name="resp"][value="${CSS.escape(v)}"]`);
      if (c) c.checked = true;
    });
  } else {
    const el = document.getElementById('campo-resp');
    if (el) el.value = valor;
  }
}

function leerRespuestaActual() {
  const q = cola[indice];
  if (!q) return '';

  if (q.tipo === 'radio') {
    const checked = document.querySelector('input[name="resp"]:checked');
    if (checked) return checked.value;
    // fallback: valor ya guardado por dictado/clic
    if (respuestas[q.id] !== undefined && respuestas[q.id] !== null && respuestas[q.id] !== '') {
      return respuestas[q.id];
    }
    return '';
  }

  if (q.tipo === 'checkbox') {
    const arr = Array.from(document.querySelectorAll('input[name="resp"]:checked')).map(c => c.value);
    if (arr.length) return arr;
    if (Array.isArray(respuestas[q.id]) && respuestas[q.id].length) return respuestas[q.id];
    return [];
  }

  const el = document.getElementById('campo-resp');
  if (el) {
    const v = String(el.value || '').trim();
    if (v !== '') return v;
  }
  if (respuestas[q.id] !== undefined && respuestas[q.id] !== null && respuestas[q.id] !== '') {
    return respuestas[q.id];
  }
  return '';
}

function etiquetaDeValor(q, valor) {
  if (valor === '' || valor === null || valor === undefined) return '(sin respuesta)';
  if (q.tipo === 'checkbox') {
    const arr = Array.isArray(valor) ? valor : [];
    if (!arr.length) return '(ninguno)';
    return arr.map(v => q.opciones.find(o => o.value === v)?.label || v).join(', ');
  }
  if (q.opciones) {
    return q.opciones.find(o => o.value === valor)?.label || valor;
  }
  return String(valor);
}

let _confirmando = false;

function confirmarYSiguiente() {
  if (_confirmando) return;
  _confirmando = true;

  try {
    // 1) Si hay dictado en curso, aplicar el texto ANTES de leer el campo
    const dictado = detenerMicrofono(true);
    const q = cola[indice];
    if (!q) {
      _confirmando = false;
      return;
    }

    if (dictado) {
      aplicarRespuestaVoz(dictado);
    }

    // 2) Leer valor del control o del fallback en respuestas
    let val = leerRespuestaActual();

    // 3) Si el campo de texto tiene algo, priorizarlo (última edición del usuario)
    const el = document.getElementById('campo-resp');
    if (el && q.tipo !== 'radio' && q.tipo !== 'checkbox') {
      const typed = String(el.value || '').trim();
      if (typed !== '') val = typed;
    }
    if (q.tipo === 'radio') {
      const checked = document.querySelector('input[name="resp"]:checked');
      if (checked) val = checked.value;
    }
    if (q.tipo === 'checkbox') {
      val = Array.from(document.querySelectorAll('input[name="resp"]:checked')).map(c => c.value);
      if (!val.length && Array.isArray(respuestas[q.id])) val = respuestas[q.id];
    }

    const vacio = (val === '' || val === null || val === undefined ||
      (Array.isArray(val) && !val.length));

    if (q.required && vacio) {
      mostrarMensaje('⚠ Este dato es obligatorio. Responda antes de continuar.');
      hablar('Este dato es obligatorio. Por favor responda antes de continuar.');
      _confirmando = false;
      return;
    }

    // 4) Guardar SIEMPRE en memoria + localStorage
    respuestas[q.id] = Array.isArray(val) ? val.slice() : val;

    if (q.id === 'num_familia') {
      const n = aplicarCantidadFamilia(val);
      respuestas[q.id] = n;
      let aviso = '✓ Guardado: ' + n + ' persona(s) en el núcleo (sin contar al socio).';
      if (n > 0) {
        aviso += ' Ahora los datos de cada una.';
      } else {
        aviso += ' Sin convivientes; pasamos a tipo de familia.';
      }
      mostrarMensaje(aviso);
    }

    const conf = etiquetaDeValor(q, respuestas[q.id]);
    try {
      guardarBorrador(true);
    } catch (e) {
      console.error('localStorage', e);
      mostrarMensaje('No se pudo guardar el borrador en el navegador, pero la respuesta quedó en memoria.');
    }

    if (q.id !== 'num_familia') {
      mostrarMensaje('✓ Guardado: ' + conf);
    }

    // 5) Avanzar a la siguiente pregunta ACTIVA (omite miembros fuera de cantidad)
    const next = buscarIndiceActivo(indice + 1, +1);
    if (next >= cola.length) {
      indice = cola.length;
      guardarBorrador(true);
      _confirmando = false;
      hablar('Respuesta guardada: ' + conf + '. Fin de la entrevista.', function () {
        irAResumen();
      });
      setTimeout(function () {
        const enResumen = document.getElementById('pantalla-resumen')?.classList.contains('activa');
        if (!enResumen && indice >= cola.length) irAResumen();
      }, 3500);
      return;
    }

    indice = next;
    guardarBorrador(true);
    _confirmando = false;

    // Anuncio al entrar al bloque de miembros del núcleo
    if (q.id === 'num_familia') {
      const n = parseInt(respuestas.num_familia, 10) || 0;
      if (n > 0) {
        hablar(
          'Registrado: ' + n + (n === 1 ? ' persona. ' : ' personas. ') +
          'Ahora preguntaré los datos de cada miembro del núcleo familiar.',
          function () { mostrarPreguntaActual(); }
        );
        return;
      }
      hablar('Sin convivientes. Pasamos al tipo de familia.', function () {
        mostrarPreguntaActual();
      });
      return;
    }
    mostrarPreguntaActual();
  } catch (err) {
    console.error(err);
    _confirmando = false;
    mostrarMensaje('Error al confirmar. Intente de nuevo.');
  }
}

function saltarPregunta() {
  const dictado = detenerMicrofono(true);
  if (dictado) aplicarRespuestaVoz(dictado);

  const q = cola[indice];
  if (!q) return;
  if (q.required) {
    hablar('No se puede omitir este campo obligatorio.');
    return;
  }
  if (respuestas[q.id] === undefined) respuestas[q.id] = '';
  if (q.id === 'num_familia') {
    const actual = respuestas.num_familia;
    if (actual === undefined || actual === null || actual === '') {
      aplicarCantidadFamilia(0);
      respuestas[q.id] = 0;
    } else {
      aplicarCantidadFamilia(actual);
    }
  }
  guardarBorrador(true);

  const next = buscarIndiceActivo(indice + 1, +1);
  if (next >= cola.length) {
    indice = cola.length;
    irAResumen();
  } else {
    indice = next;
    guardarBorrador(true);
    mostrarPreguntaActual();
  }
}

function preguntaAnterior() {
  if (indice <= 0) return;
  detenerMicrofono(false);
  detenerHabla();
  const prev = buscarIndiceActivo(indice - 1, -1);
  if (prev < 0) return;
  indice = prev;
  mostrarPreguntaActual();
}

function repetirPregunta() {
  cancelarEscuchaProgramada();
  detenerMicrofono(false);
  const q = cola[indice];
  if (q) {
    hablar(q.vozPregunta || q.pregunta, function () {
      programarEscuchaAutomatica(400);
    });
  }
}

function volverAPregunta(i) {
  if (!cola.length) cola = construirColaBase();
  let dest = typeof i === 'number' ? i : 0;
  if (!esPreguntaActiva(cola[dest])) {
    dest = buscarIndiceActivo(dest, +1);
    if (dest >= cola.length) dest = Math.max(0, buscarIndiceActivo(cola.length - 1, -1));
  }
  indice = dest;
  document.getElementById('pantalla-resumen').classList.remove('activa');
  document.getElementById('pantalla-entrevista').classList.add('activa');
  mostrarPreguntaActual();
}

function editarDesdeResumen(i) {
  editarDesdeResumenIdx(i);
}

/**
 * Editar cualquier pregunta del resumen (incluida vacía u omitida).
 * Si es de un miembro familiar fuera de rango, se activa al menos ese miembro.
 */
function editarDesdeResumenIdx(i) {
  if (!cola.length) cola = construirColaBase();
  let dest = typeof i === 'number' ? i : 0;
  if (dest < 0) dest = 0;
  if (dest >= cola.length) dest = cola.length - 1;

  const q = cola[dest];
  // Si la pregunta es de un familiar no activo, ampliar num_familia para poder editarla
  if (q && q._miembro != null) {
    const n = parseInt(respuestas.num_familia, 10) || 0;
    if (q._miembro > n) {
      respuestas.num_familia = q._miembro;
    }
  }

  indice = dest;
  document.getElementById('pantalla-resumen').classList.remove('activa');
  document.getElementById('pantalla-entrevista').classList.add('activa');
  // Forzar mostrar aunque antes estuviera "inactiva"
  if (cola[indice] && cola[indice]._miembro != null) {
    // ya ampliado num_familia
  }
  mostrarPreguntaActual();
}

// ═══════════════════════════════════════════
// MODO ADMIN (diagnóstico, URL, hoja, carpeta)
// ═══════════════════════════════════════════
const ADMIN_CLAVE = 'uncp-admin';
const ADMIN_KEY = 'fichaUNCP_modoAdmin';
let _adminClics = 0;
let _adminClicTimer = null;

function esModoAdmin() {
  try {
    return sessionStorage.getItem(ADMIN_KEY) === '1';
  } catch (e) {
    return false;
  }
}

function actualizarUIAdmin() {
  const panel = document.getElementById('panel-admin');
  if (!panel) return;
  const on = esModoAdmin();
  panel.style.display = on ? 'block' : 'none';
  // Enlaces admin actualizados
  const hoja = document.getElementById('admin-link-hoja');
  const drive = document.getElementById('admin-link-drive');
  if (hoja && typeof SHEETS_HOJA_URL === 'string') hoja.href = SHEETS_HOJA_URL;
  if (drive && typeof SHEETS_DRIVE_FOLDER_URL === 'string') drive.href = SHEETS_DRIVE_FOLDER_URL;
}

/** 5 clics rápidos en “UNCP · Ficha social” o Alt+Shift+A */
function intentoEntrarAdmin() {
  _adminClics++;
  if (_adminClicTimer) clearTimeout(_adminClicTimer);
  _adminClicTimer = setTimeout(function () { _adminClics = 0; }, 2000);
  if (_adminClics >= 5) {
    _adminClics = 0;
    pedirClaveAdmin();
  }
}

function pedirClaveAdmin() {
  const clave = window.prompt('Clave de administrador:');
  if (clave == null) return;
  if (String(clave).trim() === ADMIN_CLAVE) {
    try { sessionStorage.setItem(ADMIN_KEY, '1'); } catch (e) { /* ignore */ }
    actualizarUIAdmin();
    actualizarEstadoSheetsUI();
    hablar('Modo administrador activado.');
    mostrarMensaje('🔐 Modo admin activo. Diagnóstico, URL, hoja y carpeta visibles.');
  } else {
    hablar('Clave incorrecta.');
    alert('Clave incorrecta.');
  }
}

function salirModoAdmin() {
  try { sessionStorage.removeItem(ADMIN_KEY); } catch (e) { /* ignore */ }
  const box = document.getElementById('config-sheets-box');
  if (box) box.style.display = 'none';
  actualizarUIAdmin();
  hablar('Modo administrador desactivado.');
  mostrarMensaje('Modo admin desactivado.');
}

// ═══════════════════════════════════════════
// HELPERS DE MARCADO
// ═══════════════════════════════════════════
function r() { return respuestas; }

function chk(valor, opcion) {
  return valor === opcion ? 'X' : '  ';
}

function box(label, valor, opcion) {
  return `${label} (${chk(valor, opcion)})`;
}

function siNo(valor) {
  if (valor === 'si') return 'SÍ (  X  )   NO (     )';
  if (valor === 'no') return 'SÍ (     )   NO (  X  )';
  return 'SÍ (     )   NO (     )';
}

function fmtFecha(v) {
  if (!v) return '';
  // yyyy-mm-dd -> dd/mm/yyyy
  const m = String(v).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return `${m[3]}/${m[2]}/${m[1]}`;
  return v;
}

function labelInstruccion(v) {
  const map = {
    'sin-estudios': 'Sin estudios',
    primaria: 'Primaria incompleta',
    'primaria-completa': 'Primaria completa',
    secundaria: 'Secundaria incompleta',
    'secundaria-completa': 'Secundaria completa',
    tecnico: 'Técnico',
    universitario: 'Universitario'
  };
  return map[v] || v || '';
}

function setTxt(id, v) {
  const el = document.getElementById(id);
  if (!el) return;
  const t = v == null ? '' : String(v);
  el.textContent = t;
  // Por si hay un clon o el id se repite en captura
  try {
    document.querySelectorAll('#' + CSS.escape(id)).forEach(node => {
      node.textContent = t;
    });
  } catch (e) {
    el.textContent = t;
  }
}

// ═══════════════════════════════════════════
// RESUMEN
// ═══════════════════════════════════════════
function irAResumen() {
  cancelarEscuchaProgramada();
  detenerMicrofono(false);
  detenerHabla();
  document.getElementById('pantalla-entrevista').classList.remove('activa');
  document.getElementById('pantalla-resumen').classList.add('activa');
  actualizarBotonMicro();
  pintarResumen();
  rellenarPlantillaPDF();
  actualizarEstadoSheetsUI();
  actualizarUIAdmin();

  if (obtenerUrlSheets()) {
    setTimeout(async function () {
      try {
        const resp = await fetch(obtenerUrlSheets(), { method: 'GET', redirect: 'follow', cache: 'no-store' });
        const text = await resp.text();
        let data = null;
        try { data = JSON.parse(text); } catch (e) { /* ignore */ }
        if (!data || !data.pdfReady) {
          if (esModoAdmin()) {
            mostrarAvisoScriptViejo(true);
            setEstadoSheets('Admin: script de Google sin pdfReady. Revise la URL /exec.', true);
          }
          hablar('Entrevista finalizada. Guardando la ficha.');
        } else {
          mostrarAvisoScriptViejo(false);
          hablar('Entrevista finalizada. Guardando la ficha en la nube.');
        }
      } catch (e) {
        hablar('Entrevista finalizada. Intentando guardar.');
      }
      enviarAGoogleSheets({ silencioso: false, auto: true });
    }, 400);
  } else {
    if (esModoAdmin()) {
      setEstadoSheets('Admin: falta la URL /exec. Configure en el panel.', true);
      toggleConfigSheets(true);
    } else {
      setEstadoSheets('No se pudo preparar el guardado. Avise al personal de la UNCP.', true);
    }
    hablar('Entrevista finalizada. Si no se guarda sola, avise al personal.');
  }
}

function pintarResumen() {
  const lista = document.getElementById('lista-resumen');
  let html = '';
  if (!cola.length) cola = construirColaBase();
  // Reconstruir miembros según cantidad para poder listar/editar todos
  const nFam = parseInt(respuestas.num_familia, 10) || 0;
  // Mostrar TODAS las preguntas activas (con o sin respuesta) para poder editarlas
  cola.forEach((q, i) => {
    if (!esPreguntaActiva(q)) return;
    const val = respuestas[q.id];
    const vacio = val === undefined || val === null || val === '' ||
      (Array.isArray(val) && !val.length);
    const texto = vacio ? '(sin respuesta — pulse Editar)' : etiquetaDeValor(q, val);
    const secc = SECCIONES[q.seccion] ? `<span style="display:block;font-size:0.8rem;color:var(--muted);margin-bottom:2px;">${esc(SECCIONES[q.seccion])}</span>` : '';
    html += `<div class="resumen-item">
      ${secc}
      <strong>${esc(q.pregunta)}</strong>
      <span style="${vacio ? 'opacity:0.75;font-style:italic;' : ''}">${esc(texto)}</span>
      <button type="button" class="secundario" style="margin-top:6px;min-height:36px;padding:6px 12px;font-size:14px;"
        onclick="editarDesdeResumenIdx(${i})">Editar</button>
    </div>`;
  });
  // En admin: también listar miembros no activos por si se quiere ampliar
  if (esModoAdmin()) {
    cola.forEach((q, i) => {
      if (esPreguntaActiva(q)) return;
      if (q._miembro == null) return;
      const val = respuestas[q.id];
      const vacio = val === undefined || val === null || val === '';
      if (vacio && q._miembro > nFam) return; // no saturar
      html += `<div class="resumen-item" style="border-color:var(--info);">
        <span style="display:block;font-size:0.8rem;color:var(--info);">Admin · familiar ${q._miembro}</span>
        <strong>${esc(q.pregunta)}</strong>
        <span>${esc(vacio ? '(inactivo)' : etiquetaDeValor(q, val))}</span>
        <button type="button" class="secundario" style="margin-top:6px;min-height:36px;padding:6px 12px;font-size:14px;"
          onclick="editarDesdeResumenIdx(${i})">Editar</button>
      </div>`;
    });
  }
  if (!html) html = '<p>No hay preguntas en la cola. Inicie la entrevista.</p>';
  lista.innerHTML = html;
  actualizarUIAdmin();
}

// ═══════════════════════════════════════════
// PLANTILLA PDF (formato ficha oficial)
// ═══════════════════════════════════════════
function rellenarPlantillaPDF() {
  const x = r();
  setTxt('pdf-nombres', x.nombres);
  setTxt('pdf-dni', x.dni);
  setTxt('pdf-fecha-nac', fmtFecha(x.fecha_nac));
  setTxt('pdf-edad', x.edad);
  setTxt('pdf-sexo', `M (${chk(x.sexo, 'M')})   F (${chk(x.sexo, 'F')})`);
  setTxt('pdf-depto', x.depto_nac);
  setTxt('pdf-prov', x.prov_nac);
  setTxt('pdf-dist', x.dist_nac);
  setTxt('pdf-estado-civil', [
    box('SOLTERO/A', x.estado_civil, 'soltero'),
    box('CASADO/A', x.estado_civil, 'casado'),
    box('VIUDO/A', x.estado_civil, 'viudo'),
    box('CONVIVIENTE', x.estado_civil, 'conviviente'),
    box('DIVORCIADO/A', x.estado_civil, 'divorciado'),
    box('SEPARADO/A', x.estado_civil, 'separado')
  ].join('   '));
  setTxt('pdf-direccion', x.direccion);
  setTxt('pdf-referencia', x.referencia);
  setTxt('pdf-instruccion', labelInstruccion(x.instruccion));
  setTxt('pdf-telefono', x.telefono);
  setTxt('pdf-correo', x.correo);
  setTxt('pdf-c1-nombre', x.c1_nombre);
  setTxt('pdf-c1-par', x.c1_par);
  setTxt('pdf-c1-cel', x.c1_cel);
  setTxt('pdf-c2-nombre', x.c2_nombre);
  setTxt('pdf-c2-par', x.c2_par);
  setTxt('pdf-c2-cel', x.c2_cel);

  setTxt('pdf-padre-madre', siNo(x.es_padre_madre));
  setTxt('pdf-n-hijos', x.n_hijos);
  setTxt('pdf-n-hijos-disc', x.n_hijos_disc);
  setTxt('pdf-otros-fam-disc', x.otros_fam_disc);

  const tbody = document.getElementById('pdf-familia-body');
  if (tbody) {
    tbody.innerHTML = '';
    const n = parseInt(x.num_familia, 10) || 0;
    const filas = Math.max(n, 3);
    for (let i = 1; i <= filas; i++) {
      const tr = document.createElement('tr');
      [
        x[`fam${i}_nombre`] || '',
        x[`fam${i}_parentesco`] || '',
        x[`fam${i}_edad`] || '',
        x[`fam${i}_instruccion`] || '',
        x[`fam${i}_ocupacion`] || '',
        x[`fam${i}_aporta`] === 'si' ? 'Sí' : x[`fam${i}_aporta`] === 'no' ? 'No' : '',
        x[`fam${i}_disc_enf`] || ''
      ].forEach(c => {
        const td = document.createElement('td');
        td.textContent = c;
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    }
  }

  setTxt('pdf-tipo-familia', [
    box('UNIPERSONAL', x.tipo_familia, 'unipersonal'),
    box('NUCLEAR', x.tipo_familia, 'nuclear'),
    box('MONOPARENTAL', x.tipo_familia, 'monoparental'),
    box('EXTENSA', x.tipo_familia, 'extensa'),
    'OTROS (ESPECIFICAR): ' + (x.tipo_familia === 'otros' ? (x.tipo_familia_otro || '') : (x.tipo_familia_otro || ''))
  ].join('   '));

  setTxt('pdf-relaciones', [
    box('BUENA', x.relaciones_fam, 'buena'),
    box('REGULAR', x.relaciones_fam, 'regular'),
    box('MALA', x.relaciones_fam, 'mala')
  ].join('   '));
  setTxt('pdf-comunicacion', [
    box('BUENA', x.comunicacion_fam, 'buena'),
    box('REGULAR', x.comunicacion_fam, 'regular'),
    box('MALA', x.comunicacion_fam, 'mala')
  ].join('   '));
  setTxt('pdf-red-apoyo', x.red_apoyo);
  setTxt('pdf-apreciacion', x.apreciacion_prof);

  setTxt('pdf-tipo-disc', [
    box('CEGUERA TOTAL', x.tipo_discapacidad, 'ceguera'),
    box('BAJA VISIÓN', x.tipo_discapacidad, 'baja-vision')
  ].join('   '));
  setTxt('pdf-origen', [
    box('CONGÉNITA', x.condicion, 'congenita'),
    box('ADQUIRIDA', x.condicion, 'adquirida')
  ].join('   '));
  setTxt('pdf-edad-adq', x.edad_adquisicion);
  setTxt('pdf-certificado', x.certificado_disc === 'si' ? 'Sí' : x.certificado_disc === 'no' ? 'No' : '');
  setTxt('pdf-conadis', x.carnet_conadis === 'si' ? 'Sí' : x.carnet_conadis === 'no' ? 'No' : '');
  setTxt('pdf-atencion', siNo(x.atencion_oftal));
  setTxt('pdf-lugar-atencion', x.lugar_atencion);

  let seguroTxt = [
    box('SIS', x.seguro, 'sis'),
    box('ESSALUD', x.seguro, 'essalud'),
    box('PRIVADO', x.seguro, 'privado'),
    box('OTROS', x.seguro, 'otros')
  ].join('   ');
  if (x.seguro_otro) seguroTxt += '  ' + x.seguro_otro;
  setTxt('pdf-seguro', seguroTxt);
  setTxt('pdf-cronicas', x.enfermedades_cronicas);
  setTxt('pdf-autonomia', [
    box('INDEPENDIENTE', x.autonomia, 'independiente'),
    box('APOYO PARCIAL', x.autonomia, 'parcial'),
    box('APOYO PERMANENTE', x.autonomia, 'permanente')
  ].join('   '));

  setTxt('pdf-ingreso', [
    box('SIN INGRESOS', x.ingreso, 'sin-ingresos'),
    box('< S/500', x.ingreso, 'menos-500'),
    box('S/500-900', x.ingreso, '500-900'),
    box('S/900-1500', x.ingreso, '900-1500'),
    box('> S/1500', x.ingreso, 'mas-1500')
  ].join('   '));
  setTxt('pdf-laborando', siNo(x.laborando));
  setTxt('pdf-ocupacion', x.ocupacion);
  setTxt('pdf-profesion', x.profesion);
  setTxt('pdf-fuente', [
    box('TRAB. DEPENDIENTE', x.fuente_ingresos, 'dependiente'),
    box('TRAB. INDEPENDIENTE', x.fuente_ingresos, 'independiente'),
    box('PENSIÓN/JUBILACIÓN', x.fuente_ingresos, 'pension'),
    box('PROGRAMA SOCIAL', x.fuente_ingresos, 'programa'),
    box('APOYO FAMILIAR', x.fuente_ingresos, 'apoyo-familiar'),
    box('OTROS', x.fuente_ingresos, 'otros')
  ].join('   ') + (x.programa_social_detalle ? ' — ' + x.programa_social_detalle : ''));
  setTxt('pdf-otros-ingresos', siNo(x.otros_ingresos_hogar));
  setTxt('pdf-aportante', x.principal_aportante);
  setTxt('pdf-sisfoh', [
    box('NO TIENE', x.sisfoh, 'no-tiene'),
    box('NO SABE', x.sisfoh, 'no-sabe'),
    box('POBRE', x.sisfoh, 'pobre'),
    box('POBRE EXTREMO', x.sisfoh, 'pobre-extremo'),
    box('NO POBRE', x.sisfoh, 'no-pobre')
  ].join('   '));

  setTxt('pdf-tenencia', [
    box('PROPIA', x.tenencia, 'propia'),
    box('ALQUILADA', x.tenencia, 'alquilada'),
    box('CEDIDA', x.tenencia, 'cedida'),
    'OTRO: ' + (x.tenencia_otro || '')
  ].join('   '));
  setTxt('pdf-tipo-vivienda', [
    box('CASA INDEPENDIENTE', x.tipo_vivienda, 'casa'),
    box('DEPARTAMENTO', x.tipo_vivienda, 'departamento'),
    box('CUARTO', x.tipo_vivienda, 'cuarto'),
    box('QUINTA', x.tipo_vivienda, 'quinta'),
    box('MULTIFAMILIAR', x.tipo_vivienda, 'multifamiliar'),
    'OTRO: ' + (x.tipo_vivienda_otro || '')
  ].join('   '));
  setTxt('pdf-ambientes', x.n_ambientes);
  setTxt('pdf-distribucion', [
    box('DORMITORIO PROPIO', x.distribucion_dorm, 'propio'),
    box('COMPARTE DORMITORIO', x.distribucion_dorm, 'comparte-dorm'),
    box('COMPARTE HABITACIÓN', x.distribucion_dorm, 'comparte-hab')
  ].join('   '));
  const serv = Array.isArray(x.servicios) ? x.servicios : [];
  setTxt('pdf-servicios', [
    box('AGUA POTABLE', serv.includes('agua') ? 'agua' : '', 'agua'),
    box('DESAGÜE', serv.includes('desague') ? 'desague' : '', 'desague'),
    box('ENERGÍA ELÉCTRICA', serv.includes('luz') ? 'luz' : '', 'luz'),
    box('INTERNET', serv.includes('internet') ? 'internet' : '', 'internet')
  ].join('   '));
  setTxt('pdf-barreras', [
    box('PRESENTA BARRERAS', x.barreras, 'si'),
    box('NO PRESENTA BARRERAS', x.barreras, 'no')
  ].join('   '));
  setTxt('pdf-barreras-esp', x.barreras_especificar);
  setTxt('pdf-diagnostico', x.diagnostico_social);
  setTxt('pdf-mortuoria', x.derecho_mortuoria);
  setTxt('pdf-fecha-ent', fmtFecha(x.fecha_entrevista));
  setTxt('pdf-profesional', x.profesional);
  setTxt('pdf-firma', x.firma);
}

/**
 * Descarga PDF = misma plantilla que se imprime (#plantilla-pdf).
 * También es la que se sube a Drive al guardar la ficha.
 */
async function generarPDF() {
  rellenarPlantillaPDF();
  hablar('Generando PDF de la ficha impresa…');
  try {
    const b64 = await generarPdfVisualBase64();
    if (!b64 || b64.length < 400) {
      hablar('No se pudo armar el archivo. Abriendo impresión de la misma plantilla.');
      imprimirPlantillaPdfLocal();
      return;
    }
    const blob = base64ToBlob(b64, 'application/pdf');
    const nombre = nombreArchivoPdfFicha();
    descargarBlob(blob, nombre);
    hablar('PDF descargado: ' + nombre + '. Es la misma ficha que se imprime y se guarda en la nube.');
    mostrarMensaje('✅ PDF descargado (versión impresa): ' + nombre);
  } catch (err) {
    console.error(err);
    hablar('No se pudo generar el PDF. Intentando imprimir la plantilla.');
    imprimirPlantillaPdfLocal();
  }
}

/** Imprimir solo la plantilla oficial (sin UI, sin contenido duplicado). */
function imprimirPlantillaPdfLocal() {
  rellenarPlantillaPDF();
  const el = document.getElementById('plantilla-pdf');
  if (el) el.classList.add('exportando');
  document.body.classList.add('imprimiendo-ficha');
  hablar('Abriendo impresión de la ficha oficial. Elija Guardar como PDF si lo desea.');
  setTimeout(() => {
    window.print();
    setTimeout(() => {
      if (el) el.classList.remove('exportando');
      document.body.classList.remove('imprimiendo-ficha');
    }, 800);
  }, 350);
}

/** base64 → Blob (PDF, Excel, etc.) */
function base64ToBlob(b64, mime) {
  const raw = String(b64 || '').replace(/^data:[^;]+;base64,/i, '');
  const bin = atob(raw);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return new Blob([bytes], { type: mime || 'application/octet-stream' });
}

function arrayBufferToBase64(buffer) {
  const bytes = new Uint8Array(buffer);
  let binary = '';
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
  }
  return btoa(binary);
}

/** Nombre de archivo PDF para Drive. */
function nombreArchivoPdfFicha() {
  const x = respuestas || {};
  const nom = String(x.nombres || 'socio')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^\w\s-]/g, '')
    .trim()
    .replace(/\s+/g, '_')
    .slice(0, 40) || 'socio';
  const dni = String(x.dni || '').replace(/\D/g, '').slice(0, 12);
  const f = String(x.fecha_entrevista || new Date().toISOString().slice(0, 10));
  return 'Ficha_Social_' + (dni ? dni + '_' : '') + nom + '_' + f + '.pdf';
}

/** Logo embebido como PNG (jsPDF no siempre acepta webp). */
async function obtenerLogoDataUrl() {
  if (window._logoDataUrlFicha) return window._logoDataUrlFicha;
  try {
    const resp = await fetch('logo-blanco-uncp-scaled.webp', { cache: 'force-cache' });
    if (!resp.ok) throw new Error('logo http ' + resp.status);
    const blob = await resp.blob();
    const dataUrl = await new Promise((resolve, reject) => {
      const img = new Image();
      img.onload = function () {
        const canvas = document.createElement('canvas');
        canvas.width = img.naturalWidth || 200;
        canvas.height = img.naturalHeight || 80;
        const ctx = canvas.getContext('2d');
        ctx.fillStyle = '#000000';
        ctx.fillRect(0, 0, canvas.width, canvas.height);
        ctx.drawImage(img, 0, 0);
        resolve(canvas.toDataURL('image/png'));
      };
      img.onerror = reject;
      img.src = URL.createObjectURL(blob);
    });
    window._logoDataUrlFicha = dataUrl;
    return dataUrl;
  } catch (e) {
    console.warn('Logo no embebido', e);
    return '';
  }
}

function pdfTxt(v) {
  if (v == null || v === undefined) return '';
  if (Array.isArray(v)) return v.join(', ');
  return String(v);
}

function pdfMarca(valor, opcion) {
  return valor === opcion ? 'X' : ' ';
}

function pdfOpc(label, valor, opcion) {
  return label + ' (' + pdfMarca(valor, opcion) + ')';
}

function pdfSiNo(valor) {
  if (valor === 'si') return 'SI ( X )   NO (   )';
  if (valor === 'no') return 'SI (   )   NO ( X )';
  return 'SI (   )   NO (   )';
}

/**
 * PDF = MISMA plantilla que se imprime (#plantilla-pdf).
 * Captura HTML → páginas A4 (html2canvas + jsPDF).
 * Usado en: descarga local y subida a Google Drive.
 * @returns {Promise<string|null>} base64 sin prefijo
 */
async function generarPdfVisualBase64() {
  rellenarPlantillaPDF();
  const el = document.getElementById('plantilla-pdf');
  if (!el) {
    console.warn('No existe #plantilla-pdf');
    return generarPdfTextoJsPdfBase64();
  }

  const jspdfNS = window.jspdf;
  if (!jspdfNS || !jspdfNS.jsPDF) {
    console.warn('jsPDF no cargó');
    mostrarMensaje('⚠ No se cargó jsPDF. Revise internet y recargue la página.');
    return null;
  }
  if (typeof html2canvas !== 'function') {
    console.warn('html2canvas no cargó — respaldo tablas jsPDF');
    return generarPdfTextoJsPdfBase64();
  }

  const { jsPDF } = jspdfNS;
  const prevClass = el.className;
  el.classList.add('exportando', 'capturando');
  el.setAttribute('aria-hidden', 'false');

  // Esperar layout + logo
  await new Promise(function (r) { setTimeout(r, 120); });
  try {
    const imgs = el.querySelectorAll('img');
    await Promise.all(Array.prototype.map.call(imgs, function (img) {
      if (img.complete) return Promise.resolve();
      return new Promise(function (res) {
        img.onload = res;
        img.onerror = res;
        setTimeout(res, 1500);
      });
    }));
  } catch (eImg) { /* ignore */ }

  let canvas;
  try {
    canvas = await html2canvas(el, {
      scale: 2,
      useCORS: true,
      allowTaint: true,
      backgroundColor: '#ffffff',
      logging: false,
      scrollX: 0,
      scrollY: 0,
      windowWidth: el.scrollWidth || 794,
      windowHeight: el.scrollHeight || el.offsetHeight
    });
  } catch (eCap) {
    console.error('html2canvas', eCap);
    el.className = prevClass;
    el.setAttribute('aria-hidden', 'true');
    return generarPdfTextoJsPdfBase64();
  }

  el.className = prevClass;
  el.classList.remove('exportando', 'capturando');
  el.setAttribute('aria-hidden', 'true');

  if (!canvas || canvas.width < 10 || canvas.height < 10) {
    console.warn('Canvas vacío');
    return generarPdfTextoJsPdfBase64();
  }

  try {
    // Calidad un poco menor = archivo más liviano (evita fallos de red al subir)
    const imgData = canvas.toDataURL('image/jpeg', 0.78);
    const pdf = new jsPDF({ unit: 'mm', format: 'a4', orientation: 'portrait' });
    const pageW = pdf.internal.pageSize.getWidth();
    const pageH = pdf.internal.pageSize.getHeight();
    // Ancho de página con pequeño margen
    const margin = 0;
    const imgW = pageW - margin * 2;
    const imgH = (canvas.height * imgW) / canvas.width;

    let heightLeft = imgH;
    let y = margin;
    pdf.addImage(imgData, 'JPEG', margin, y, imgW, imgH);
    heightLeft -= pageH;

    while (heightLeft > 1) {
      y = heightLeft - imgH + margin;
      pdf.addPage();
      pdf.addImage(imgData, 'JPEG', margin, y, imgW, imgH);
      heightLeft -= pageH;
    }

    const dataUri = pdf.output('datauristring');
    const b64 = (dataUri && dataUri.indexOf(',') >= 0) ? dataUri.split(',')[1] : '';
    if (!b64 || b64.length < 400) {
      console.warn('PDF plantilla base64 vacío');
      return generarPdfTextoJsPdfBase64();
    }
    console.log(
      'PDF OK plantilla impresa, len=', b64.length,
      'páginas≈', Math.ceil(imgH / pageH),
      'nombres=', (respuestas && respuestas.nombres) || ''
    );
    return b64;
  } catch (err) {
    console.error('PDF plantilla output', err);
    return generarPdfTextoJsPdfBase64();
  }
}

/**
 * Respaldo: PDF de tablas con jsPDF (solo si falla la captura de la plantilla impresa).
 * @returns {Promise<string|null>}
 */
async function generarPdfTextoJsPdfBase64() {
  const jspdfNS = window.jspdf;
  if (!jspdfNS || !jspdfNS.jsPDF) return null;
  const { jsPDF } = jspdfNS;
  if (typeof jsPDF.API.autoTable !== 'function') return null;

  const x = respuestas || {};
  const logoUrl = await obtenerLogoDataUrl();
  const ML = 14;
  const MR = 14;
  const MT = 12;
  const MB = 12;
  const doc = new jsPDF({ unit: 'mm', format: 'a4', orientation: 'portrait' });
  const pageW = doc.internal.pageSize.getWidth();
  const pageH = doc.internal.pageSize.getHeight();
  const usableW = pageW - ML - MR;
  const gray = [235, 235, 235];
  let cursorY = MT;

  function checkPage(need) {
    if (cursorY + need > pageH - MB) {
      doc.addPage();
      cursorY = MT;
    }
  }
  function seccion(titulo) {
    checkPage(8);
    cursorY += 2;
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(9);
    doc.setTextColor(80, 80, 80);
    doc.text(String(titulo), ML, cursorY);
    doc.setTextColor(0, 0, 0);
    cursorY += 3;
  }
  function filas(pares) {
    if (!pares || !pares.length) return;
    checkPage(10);
    const body = pares.map(function (p) {
      return [
        { content: pdfTxt(p[0]), styles: { fontStyle: 'bold', fillColor: gray, fontSize: 8, cellWidth: usableW * 0.32 } },
        { content: pdfTxt(p[1]), styles: { fontSize: 9, cellWidth: usableW * 0.68 } }
      ];
    });
    doc.autoTable({
      startY: cursorY,
      margin: { left: ML, right: MR, bottom: MB },
      tableWidth: usableW,
      theme: 'grid',
      showHead: false,
      body: body,
      styles: {
        textColor: [0, 0, 0], lineColor: [0, 0, 0], lineWidth: 0.3,
        cellPadding: 2, valign: 'middle', overflow: 'linebreak', minCellHeight: 6
      }
    });
    cursorY = (doc.lastAutoTable && doc.lastAutoTable.finalY) ? doc.lastAutoTable.finalY + 2 : cursorY + 8;
  }

  doc.setDrawColor(0);
  doc.setLineWidth(0.6);
  doc.rect(ML, cursorY, usableW, 20);
  if (logoUrl) {
    try {
      const logoW = 26;
      const logoH = 9;
      doc.setFillColor(0, 0, 0);
      doc.rect(pageW / 2 - logoW / 2 - 1.5, cursorY + 1.5, logoW + 3, logoH + 1.5, 'F');
      doc.addImage(logoUrl, 'PNG', pageW / 2 - logoW / 2, cursorY + 2, logoW, logoH);
    } catch (eLogo) { /* ignore */ }
  }
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(13);
  doc.text('FICHA SOCIAL UNCP', pageW / 2, cursorY + 14.5, { align: 'center' });
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(8);
  doc.text('Union Nacional de Ciegos del Peru - Formato oficial', pageW / 2, cursorY + 18, { align: 'center' });
  cursorY += 24;

  seccion('I. DATOS GENERALES DEL SOCIO(A)');
  filas([
    ['APELLIDOS Y NOMBRES', x.nombres],
    ['N° DNI / CARNET', x.dni],
    ['FECHA DE NACIMIENTO', fmtFecha(x.fecha_nac)],
    ['EDAD', x.edad],
    ['SEXO', 'M (' + pdfMarca(x.sexo, 'M') + ')    F (' + pdfMarca(x.sexo, 'F') + ')'],
    ['DEPARTAMENTO NAC.', x.depto_nac],
    ['PROVINCIA NAC.', x.prov_nac],
    ['DISTRITO NAC.', x.dist_nac],
    ['ESTADO CIVIL', [
      pdfOpc('SOLTERO/A', x.estado_civil, 'soltero'),
      pdfOpc('CASADO/A', x.estado_civil, 'casado'),
      pdfOpc('VIUDO/A', x.estado_civil, 'viudo'),
      pdfOpc('CONVIVIENTE', x.estado_civil, 'conviviente'),
      pdfOpc('DIVORCIADO/A', x.estado_civil, 'divorciado'),
      pdfOpc('SEPARADO/A', x.estado_civil, 'separado')
    ].join('  ')],
    ['DIRECCION ACTUAL', x.direccion],
    ['REFERENCIA', x.referencia],
    ['NIVEL EDUCATIVO', labelInstruccion(x.instruccion)],
    ['CELULAR', x.telefono],
    ['CORREO ELECTRONICO', x.correo],
    ['CONTACTO 1 - NOMBRE', x.c1_nombre],
    ['CONTACTO 1 - PARENTESCO', x.c1_par],
    ['CONTACTO 1 - CELULAR', x.c1_cel],
    ['CONTACTO 2 - NOMBRE', x.c2_nombre],
    ['CONTACTO 2 - PARENTESCO', x.c2_par],
    ['CONTACTO 2 - CELULAR', x.c2_cel]
  ]);

  seccion('II. CARACTERIZACION FAMILIAR');
  filas([
    ['¿ES USTED PADRE O MADRE?', pdfSiNo(x.es_padre_madre)],
    ['N° TOTAL DE HIJOS', x.n_hijos],
    ['N° HIJOS CON DISC. VISUAL', x.n_hijos_disc],
    ['OTROS FAM. DISC. VISUAL', x.otros_fam_disc]
  ]);

  seccion('III. COMPOSICION DEL NUCLEO FAMILIAR');
  const nFam = parseInt(x.num_familia, 10) || 0;
  filas([['N° CONVIVIENTES (sin el socio)', nFam]]);
  for (let i = 1; i <= Math.max(nFam, 1); i++) {
    if (!x['fam' + i + '_nombre'] && i > nFam) continue;
    filas([
      ['FAMILIAR ' + i + ' - NOMBRES', x['fam' + i + '_nombre']],
      ['FAMILIAR ' + i + ' - PARENTESCO', x['fam' + i + '_parentesco']],
      ['FAMILIAR ' + i + ' - EDAD', x['fam' + i + '_edad']],
      ['FAMILIAR ' + i + ' - NIVEL EDUCATIVO', x['fam' + i + '_instruccion']],
      ['FAMILIAR ' + i + ' - OCUPACION', x['fam' + i + '_ocupacion']],
      ['FAMILIAR ' + i + ' - APORTA INGRESOS', x['fam' + i + '_aporta'] === 'si' ? 'Si' : x['fam' + i + '_aporta'] === 'no' ? 'No' : ''],
      ['FAMILIAR ' + i + ' - DISC./ENFERMEDAD', x['fam' + i + '_disc_enf']]
    ]);
  }
  filas([['TIPO DE FAMILIA', [
    pdfOpc('UNIPERSONAL', x.tipo_familia, 'unipersonal'),
    pdfOpc('NUCLEAR', x.tipo_familia, 'nuclear'),
    pdfOpc('MONOPARENTAL', x.tipo_familia, 'monoparental'),
    pdfOpc('EXTENSA', x.tipo_familia, 'extensa'),
    'OTROS: ' + pdfTxt(x.tipo_familia_otro)
  ].join('  ')]]);

  seccion('IV. APRECIACION DEL NUCLEO FAMILIAR');
  filas([
    ['RELACIONES FAMILIARES', [pdfOpc('BUENA', x.relaciones_fam, 'buena'), pdfOpc('REGULAR', x.relaciones_fam, 'regular'), pdfOpc('MALA', x.relaciones_fam, 'mala')].join('  ')],
    ['COMUNICACION FAMILIAR', [pdfOpc('BUENA', x.comunicacion_fam, 'buena'), pdfOpc('REGULAR', x.comunicacion_fam, 'regular'), pdfOpc('MALA', x.comunicacion_fam, 'mala')].join('  ')],
    ['PRINCIPAL RED DE APOYO', x.red_apoyo],
    ['APRECIACION PROFESIONAL', x.apreciacion_prof]
  ]);

  seccion('V. CONDICION DE DISCAPACIDAD VISUAL');
  filas([
    ['CONDICION', [pdfOpc('CEGUERA TOTAL', x.tipo_discapacidad, 'ceguera'), pdfOpc('BAJA VISION', x.tipo_discapacidad, 'baja-vision')].join('  ')],
    ['ORIGEN', [pdfOpc('CONGENITA', x.condicion, 'congenita'), pdfOpc('ADQUIRIDA', x.condicion, 'adquirida')].join('  ')],
    ['EDAD DE ADQUISICION', x.edad_adquisicion],
    ['CERTIFICADO DE DISCAPACIDAD', x.certificado_disc === 'si' ? 'Si' : x.certificado_disc === 'no' ? 'No' : ''],
    ['CARNET CONADIS', x.carnet_conadis === 'si' ? 'Si' : x.carnet_conadis === 'no' ? 'No' : ''],
    ['ATENCION OFTALMOLOGICA', pdfSiNo(x.atencion_oftal)],
    ['LUGAR / INSTITUCION', x.lugar_atencion]
  ]);

  seccion('VI. ESTADO DE SALUD');
  filas([
    ['SEGURO DE SALUD', [
      pdfOpc('SIS', x.seguro, 'sis'), pdfOpc('ESSALUD', x.seguro, 'essalud'),
      pdfOpc('PRIVADO', x.seguro, 'privado'), pdfOpc('OTROS', x.seguro, 'otros')
    ].join('  ') + (x.seguro_otro ? ' ' + x.seguro_otro : '')],
    ['ENFERMEDADES CRONICAS', x.enfermedades_cronicas],
    ['AUTONOMIA', [
      pdfOpc('INDEPENDIENTE', x.autonomia, 'independiente'),
      pdfOpc('APOYO PARCIAL', x.autonomia, 'parcial'),
      pdfOpc('APOYO PERMANENTE', x.autonomia, 'permanente')
    ].join('  ')]
  ]);

  const serv = Array.isArray(x.servicios) ? x.servicios : [];
  seccion('VII. SITUACION LABORAL Y ECONOMICA');
  filas([
    ['INGRESO MENSUAL', [
      pdfOpc('SIN INGRESOS', x.ingreso, 'sin-ingresos'), pdfOpc('< S/500', x.ingreso, 'menos-500'),
      pdfOpc('S/500-900', x.ingreso, '500-900'), pdfOpc('S/900-1500', x.ingreso, '900-1500'),
      pdfOpc('> S/1500', x.ingreso, 'mas-1500')
    ].join('  ')],
    ['¿LABORA ACTUALMENTE?', pdfSiNo(x.laborando)],
    ['OCUPACION', x.ocupacion],
    ['PROFESION U OFICIO', x.profesion],
    ['FUENTE PRINCIPAL DE INGRESOS', [
      pdfOpc('TRAB. DEPENDIENTE', x.fuente_ingresos, 'dependiente'),
      pdfOpc('TRAB. INDEPENDIENTE', x.fuente_ingresos, 'independiente'),
      pdfOpc('PENSION', x.fuente_ingresos, 'pension'),
      pdfOpc('PROGRAMA SOCIAL', x.fuente_ingresos, 'programa'),
      pdfOpc('APOYO FAMILIAR', x.fuente_ingresos, 'apoyo-familiar'),
      pdfOpc('OTROS', x.fuente_ingresos, 'otros')
    ].join('  ') + (x.programa_social_detalle ? ' - ' + x.programa_social_detalle : '')],
    ['¿OTROS INGRESOS EN EL HOGAR?', pdfSiNo(x.otros_ingresos_hogar)],
    ['PRINCIPAL APORTANTE ECONOMICO', x.principal_aportante],
    ['CLASIFICACION SISFOH', [
      pdfOpc('NO TIENE', x.sisfoh, 'no-tiene'), pdfOpc('NO SABE', x.sisfoh, 'no-sabe'),
      pdfOpc('POBRE', x.sisfoh, 'pobre'), pdfOpc('POBRE EXTREMO', x.sisfoh, 'pobre-extremo'),
      pdfOpc('NO POBRE', x.sisfoh, 'no-pobre')
    ].join('  ')]
  ]);

  seccion('VIII. CONDICION DE VIVIENDA');
  filas([
    ['TENENCIA DE LA VIVIENDA', [
      pdfOpc('PROPIA', x.tenencia, 'propia'), pdfOpc('ALQUILADA', x.tenencia, 'alquilada'),
      pdfOpc('CEDIDA', x.tenencia, 'cedida'), 'OTRO: ' + pdfTxt(x.tenencia_otro)
    ].join('  ')],
    ['TIPO DE VIVIENDA', [
      pdfOpc('CASA INDEPENDIENTE', x.tipo_vivienda, 'casa'), pdfOpc('DEPARTAMENTO', x.tipo_vivienda, 'departamento'),
      pdfOpc('CUARTO', x.tipo_vivienda, 'cuarto'), pdfOpc('QUINTA', x.tipo_vivienda, 'quinta'),
      pdfOpc('MULTIFAMILIAR', x.tipo_vivienda, 'multifamiliar'), 'OTRO: ' + pdfTxt(x.tipo_vivienda_otro)
    ].join('  ')],
    ['N° AMBIENTES (dormir)', x.n_ambientes],
    ['DISTRIBUCION', [
      pdfOpc('DORM. PROPIO', x.distribucion_dorm, 'propio'),
      pdfOpc('COMPARTE DORM.', x.distribucion_dorm, 'comparte-dorm'),
      pdfOpc('COMPARTE HAB.', x.distribucion_dorm, 'comparte-hab')
    ].join('  ')],
    ['SERVICIOS BASICOS', [
      pdfOpc('AGUA', serv.includes('agua') ? 'agua' : '', 'agua'),
      pdfOpc('DESAGUE', serv.includes('desague') ? 'desague' : '', 'desague'),
      pdfOpc('LUZ', serv.includes('luz') ? 'luz' : '', 'luz'),
      pdfOpc('INTERNET', serv.includes('internet') ? 'internet' : '', 'internet')
    ].join('  ')],
    ['ACCESIBILIDAD', [pdfOpc('PRESENTA BARRERAS', x.barreras, 'si'), pdfOpc('NO PRESENTA', x.barreras, 'no')].join('  ')],
    ['BARRERAS (ESPECIFICAR)', x.barreras_especificar],
    ['DIAGNOSTICO SOCIAL', x.diagnostico_social],
    ['DERECHO DE MORTUORIA', x.derecho_mortuoria]
  ]);

  seccion('REGISTRO ADMINISTRATIVO');
  filas([
    ['FECHA DE ENTREVISTA', fmtFecha(x.fecha_entrevista)],
    ['PROFESIONAL RESPONSABLE', x.profesional],
    ['FIRMA O HUELLA DEL SOCIO', x.firma || '']
  ]);

  try {
    const dataUri = doc.output('datauristring');
    const b64 = (dataUri && dataUri.indexOf(',') >= 0) ? dataUri.split(',')[1] : '';
    return (b64 && b64.length > 400) ? b64 : null;
  } catch (err) {
    console.error('PDF texto fallback', err);
    return null;
  }
}

// ═══════════════════════════════════════════
// EXCEL (formato Ficha social JULIO 2026)
// ═══════════════════════════════════════════
function b64ToArrayBuffer(b64) {
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes.buffer;
}

/** Extrae texto de una celda ExcelJS (string, número o richText). */
function cellText(val) {
  if (val == null || val === '') return '';
  if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') return String(val);
  if (typeof val === 'object') {
    if (Array.isArray(val.richText)) return val.richText.map(p => p.text || '').join('');
    if (val.text != null) return String(val.text);
    if (val.result != null) return String(val.result);
  }
  return String(val);
}

/** Escribe valor en celda de respuesta (texto azul negrita). No borra si viene vacío. */
function setCell(ws, addr, value) {
  if (value === undefined || value === null || value === '') return;
  const cell = ws.getCell(addr);
  cell.value = value;
  const prev = cell.font || {};
  cell.font = {
    name: prev.name || 'Calibri',
    size: prev.size || 12,
    bold: true,
    color: { argb: 'FF0000FF' },
    italic: prev.italic || false
  };
}

/** Marca o desmarca el primer paréntesis de la celda (conserva la etiqueta de la plantilla). */
function markCheck(ws, addr, selected, fallbackLabel) {
  const cell = ws.getCell(addr);
  let text = cellText(cell.value);
  if (!text && fallbackLabel) text = fallbackLabel;
  if (!text) return;
  if (/\([^)]*\)/.test(text)) {
    text = text.replace(/\([^)]*\)/, selected ? '(  X  )' : '(     )');
  } else if (selected) {
    text = text.replace(/\s*$/, '') + ' (  X  )';
  }
  cell.value = text;
}

/** Marca un mapa value->celda según la respuesta seleccionada. */
function markOptionMap(ws, valueToAddr, selectedValue, fallbacks) {
  Object.entries(valueToAddr).forEach(([val, addr]) => {
    const fb = fallbacks && fallbacks[addr] ? fallbacks[addr] : '';
    markCheck(ws, addr, selectedValue === val, fb);
  });
}

/** Marca SI/NO en una celda que contiene ambos paréntesis (p. ej. C40, C42). */
function markSiNoCell(ws, addr, valor, fallback) {
  const cell = ws.getCell(addr);
  let text = cellText(cell.value) || fallback || '';
  if (!text) return;
  // Reemplaza el 1er ( ) = SI y el 2º = NO
  let n = 0;
  text = text.replace(/\([^)]*\)/g, () => {
    n += 1;
    if (valor === 'si') return n === 1 ? '(  X  )' : '(      )';
    if (valor === 'no') return n === 1 ? '(      )' : '(  X  )';
    return n === 1 ? '(      )' : '(      )';
  });
  cell.value = text;
}

/**
 * Genera el Excel de la ficha (misma data que el PDF, plantilla oficial).
 * @returns {Promise<{ base64: string, nombre: string, blob: Blob }|null>}
 */
async function generarExcelBuffer() {
  if (typeof ExcelJS === 'undefined') {
    throw new Error('Falta exceljs.min.js');
  }
  const wb = new ExcelJS.Workbook();
  if (typeof PLANTILLA_FICHA_B64 !== 'undefined' && PLANTILLA_FICHA_B64) {
    await wb.xlsx.load(b64ToArrayBuffer(PLANTILLA_FICHA_B64));
  } else {
    const resp = await fetch('plantilla_ficha.xlsx');
    if (!resp.ok) throw new Error('No se pudo cargar plantilla_ficha.xlsx');
    await wb.xlsx.load(await resp.arrayBuffer());
  }
  rellenarLibroExcel(wb);
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });
  const nombre = nombreArchivoFicha('xlsx');
  const base64 = arrayBufferToBase64(buffer);
  return { base64, nombre, blob };
}

async function generarExcel() {
  try {
    if (typeof ExcelJS === 'undefined') {
      hablar('No se cargó la librería Excel. Verifique que exceljs.min.js esté en la misma carpeta.');
      alert('Falta exceljs.min.js en la carpeta uciego.');
      return;
    }
    hablar('Generando Excel con la misma información del PDF…');
    const out = await generarExcelBuffer();
    if (!out) throw new Error('Excel vacío');
    descargarBlob(out.blob, out.nombre);
    hablar('Excel descargado: ' + out.nombre);
    mostrarMensaje('✅ Excel descargado: ' + out.nombre);
  } catch (err) {
    console.error(err);
    hablar('Error al generar el Excel. Intente de nuevo.');
    alert('Error al generar Excel: ' + (err && err.message ? err.message : err));
  }
}

/** Base64 del Excel (para subir a Drive junto con el PDF). */
async function generarExcelBase64() {
  try {
    const out = await generarExcelBuffer();
    return out && out.base64 ? out : null;
  } catch (e) {
    console.warn('Excel base64', e);
    return null;
  }
}

function rellenarLibroExcel(wb) {
  const ws = wb.getWorksheet('Ficha Social') || wb.worksheets[0];
  if (!ws) throw new Error('No se encontró la hoja «Ficha Social» en la plantilla.');
  rellenarHojaExcel(ws);
  const reg = wb.getWorksheet('Registro Socios');
  if (reg) rellenarRegistroSocios(reg);
}

function siNoTxt(v) {
  if (v === 'si') return 'Sí';
  if (v === 'no') return 'No';
  return v || '';
}

function rellenarHojaExcel(ws) {
  const x = r();

  // —— I. Datos generales ——
  setCell(ws, 'C3', x.nombres);
  setCell(ws, 'C4', x.dni);
  setCell(ws, 'E4', fmtFecha(x.fecha_nac));
  if (x.edad != null && x.edad !== '') setCell(ws, 'G4', x.edad);
  // Sexo: conserva etiqueta SEXO y marca M/F
  {
    const cell = ws.getCell('H4');
    let t = cellText(cell.value) || 'SEXO\nM (    )   F (   )';
    // Dos casillas: M luego F
    let n = 0;
    t = t.replace(/\([^)]*\)/g, () => {
      n += 1;
      if (n === 1) return x.sexo === 'M' ? '(  X  )' : '(    )';
      if (n === 2) return x.sexo === 'F' ? '(  X  )' : '(   )';
      return '(     )';
    });
    cell.value = t;
  }

  setCell(ws, 'D5', x.depto_nac);
  setCell(ws, 'F5', x.prov_nac);
  setCell(ws, 'H5', x.dist_nac);

  markOptionMap(ws, {
    soltero: 'C6', casado: 'D6', viudo: 'E6',
    conviviente: 'F6', divorciado: 'G6', separado: 'H6'
  }, x.estado_civil, {
    C6: 'SOLTERO/A  (     )',
    D6: 'CASADO/A  (     )',
    E6: 'VIUDO/A  (     )',
    F6: 'CONVIVIENTE (     )',
    G6: 'DIVORCIADO/A \n(     )',
    H6: 'SEPARADO/A \n(     )'
  });

  setCell(ws, 'C7', x.direccion);
  setCell(ws, 'C8', x.referencia);
  setCell(ws, 'H8', labelInstruccion(x.instruccion));
  setCell(ws, 'C9', x.telefono);
  setCell(ws, 'F9', x.correo);

  setCell(ws, 'D10', x.c1_nombre);
  setCell(ws, 'F10', x.c1_par);
  setCell(ws, 'H10', x.c1_cel);
  setCell(ws, 'D11', x.c2_nombre);
  setCell(ws, 'F11', x.c2_par);
  setCell(ws, 'H11', x.c2_cel);

  // —— II. Caracterización familiar ——
  setCell(ws, 'C13', siNoTxt(x.es_padre_madre));
  if (x.n_hijos != null && x.n_hijos !== '') setCell(ws, 'F13', x.n_hijos);
  if (x.n_hijos_disc != null && x.n_hijos_disc !== '') setCell(ws, 'H13', x.n_hijos_disc);
  setCell(ws, 'E14', x.otros_fam_disc);

  // —— III. Composición familiar (filas 17–23) ——
  const nFam = parseInt(x.num_familia, 10) || 0;
  for (let i = 1; i <= 7; i++) {
    const row = 16 + i;
    if (i <= nFam || x[`fam${i}_nombre`]) {
      setCell(ws, `B${row}`, x[`fam${i}_nombre`]);
      setCell(ws, `C${row}`, x[`fam${i}_parentesco`]);
      if (x[`fam${i}_edad`] != null && x[`fam${i}_edad`] !== '') setCell(ws, `D${row}`, x[`fam${i}_edad`]);
      setCell(ws, `E${row}`, x[`fam${i}_instruccion`]);
      setCell(ws, `F${row}`, x[`fam${i}_ocupacion`]);
      setCell(ws, `G${row}`, siNoTxt(x[`fam${i}_aporta`]));
      setCell(ws, `H${row}`, x[`fam${i}_disc_enf`]);
    }
  }

  markOptionMap(ws, {
    unipersonal: 'C24', nuclear: 'D24', monoparental: 'E24', extensa: 'F24'
  }, x.tipo_familia, {
    C24: 'UNIPERSONAL (  )',
    D24: 'NUCLEAR (  )',
    E24: 'MONOPARENTAL (  )',
    F24: 'EXTENSA (  )'
  });
  {
    const g24 = ws.getCell('G24');
    let t = cellText(g24.value) || 'OTROS (ESPECIFICAR): ';
    const detalle = x.tipo_familia_otro || (x.tipo_familia === 'otros' ? 'X' : '');
    if (/OTROS/i.test(t)) {
      t = 'OTROS (ESPECIFICAR): ' + (detalle || '');
    } else if (detalle) {
      t = t + ' ' + detalle;
    }
    g24.value = t;
  }

  // —— IV. Apreciación ——
  markOptionMap(ws, { buena: 'C26', regular: 'E26', mala: 'G26' }, x.relaciones_fam, {
    C26: 'BUENA (      )', E26: 'REGULAR (      )', G26: 'MALA (      )'
  });
  markOptionMap(ws, { buena: 'C27', regular: 'E27', mala: 'G27' }, x.comunicacion_fam, {
    C27: 'BUENA (      )', E27: 'REGULAR (      )', G27: 'MALA (      )'
  });
  setCell(ws, 'E28', x.red_apoyo);
  setCell(ws, 'E29', x.apreciacion_prof);

  // —— V. Discapacidad visual ——
  markCheck(ws, 'C31', x.tipo_discapacidad === 'ceguera', 'CEGUERA TOTAL (    )');
  markCheck(ws, 'D31', x.tipo_discapacidad === 'baja-vision', 'BAJA VISIÓN (     )');
  markCheck(ws, 'F31', x.condicion === 'congenita', 'CONGÉNITA (     )');
  markCheck(ws, 'G31', x.condicion === 'adquirida', 'ADQUIRIDA (     )');
  if (x.edad_adquisicion != null && x.edad_adquisicion !== '') setCell(ws, 'D32', x.edad_adquisicion);
  setCell(ws, 'F32', siNoTxt(x.certificado_disc));
  setCell(ws, 'H32', siNoTxt(x.carnet_conadis));
  markCheck(ws, 'D33', x.atencion_oftal === 'si', 'SI (     )');
  markCheck(ws, 'E33', x.atencion_oftal === 'no', 'NO (     )');
  setCell(ws, 'G33', x.lugar_atencion);

  // —— VI. Salud ——
  markOptionMap(ws, { sis: 'C35', essalud: 'D35', privado: 'E35', otros: 'F35' }, x.seguro, {
    C35: 'SIS (     )', D35: 'ESSALUD (     )', E35: 'PRIVADO (     )', F35: 'OTROS (      )'
  });
  if (x.seguro === 'otros' && x.seguro_otro) {
    const c = ws.getCell('F35');
    c.value = cellText(c.value) + ' ' + x.seguro_otro;
  }
  setCell(ws, 'D36', x.enfermedades_cronicas);
  markCheck(ws, 'D37', x.autonomia === 'independiente', 'INDEPENDIENTE (     )');
  markCheck(ws, 'E37', x.autonomia === 'parcial', 'REQUIERE APOYO PARCIAL (     )');
  markCheck(ws, 'G37', x.autonomia === 'permanente', 'REQUIERE APOYO PERMANENTE (     )');

  // —— VII. Laboral ——
  markOptionMap(ws, {
    'sin-ingresos': 'D39', 'menos-500': 'E39', '500-900': 'F39',
    '900-1500': 'G39', 'mas-1500': 'H39'
  }, x.ingreso, {
    D39: 'SIN INGRESOS   (     )',
    E39: '< S/500    (     )',
    F39: 'ENTRE S/500 A S/900 (     )',
    G39: 'ENTRE S/900 A S/1500 (     )',
    H39: '> S/1500 (     )'
  });

  markSiNoCell(ws, 'C40', x.laborando, '          SI (      )      NO (      )');
  setCell(ws, 'E40', x.ocupacion);
  setCell(ws, 'G40', x.profesion);

  markOptionMap(ws, {
    dependiente: 'C41', independiente: 'D41', pension: 'E41',
    programa: 'F41', 'apoyo-familiar': 'G41', otros: 'H41'
  }, x.fuente_ingresos, {
    C41: 'TRAB. DEPENDIENTE\n(      )',
    D41: 'TRAB. INDEPENDIENTE\n(      )',
    E41: 'PENSIÓN /JUBILACIÓN \n(      )',
    F41: 'PROGRAMA SOCIAL\n(      )\n___________________',
    G41: 'APOYO FAMILIAR \n(      )',
    H41: 'OTROS (      )\n____________'
  });
  if (x.programa_social_detalle && (x.fuente_ingresos === 'programa' || x.fuente_ingresos === 'otros')) {
    const addr = x.fuente_ingresos === 'programa' ? 'F41' : 'H41';
    const c = ws.getCell(addr);
    let t = cellText(c.value);
    if (/_+/.test(t)) t = t.replace(/_+/, x.programa_social_detalle);
    else t = t + '\n' + x.programa_social_detalle;
    c.value = t;
  }

  markSiNoCell(ws, 'C42', x.otros_ingresos_hogar, '      SI (      )      NO (      )');
  setCell(ws, 'F42', x.principal_aportante);

  markOptionMap(ws, {
    'no-tiene': 'D43', 'no-sabe': 'E43', pobre: 'F43',
    'pobre-extremo': 'G43', 'no-pobre': 'H43'
  }, x.sisfoh, {
    D43: 'NO TIENE (     )', E43: 'NO SABE (     )', F43: 'POBRE (     )',
    G43: 'POBRE EXTREMO (     )', H43: 'NO POBRE (     )'
  });

  // —— VIII. Vivienda ——
  markOptionMap(ws, { propia: 'C45', alquilada: 'D45', cedida: 'E45', otro: 'F45' }, x.tenencia, {
    C45: 'PROPIA (     )', D45: 'ALQUILADA (     )', E45: 'CEDIDA (     )',
    F45: 'OTRO (     ) ____________________________________'
  });
  if (x.tenencia === 'otro' && x.tenencia_otro) {
    const c = ws.getCell('F45');
    let t = cellText(c.value);
    c.value = /_+/.test(t) ? t.replace(/_+/, x.tenencia_otro) : t + ' ' + x.tenencia_otro;
  }

  markOptionMap(ws, {
    casa: 'C46', departamento: 'D46', cuarto: 'E46',
    quinta: 'F46', multifamiliar: 'G46', otro: 'H46'
  }, x.tipo_vivienda, {
    C46: 'CASA INDEPENDIENTE (     )',
    D46: 'DEPARTAMENTO (     )',
    E46: 'CUARTO (     )',
    F46: 'QUINTA \n(     )',
    G46: 'MULTIFAMILIAR (     )',
    H46: 'OTRO (      )\n____________'
  });
  if (x.tipo_vivienda === 'otro' && x.tipo_vivienda_otro) {
    const c = ws.getCell('H46');
    let t = cellText(c.value);
    c.value = /_+/.test(t) ? t.replace(/_+/, x.tipo_vivienda_otro) : t + '\n' + x.tipo_vivienda_otro;
  }

  if (x.n_ambientes != null && x.n_ambientes !== '') setCell(ws, 'C47', x.n_ambientes);
  markOptionMap(ws, {
    propio: 'E47', 'comparte-dorm': 'F47', 'comparte-hab': 'G47'
  }, x.distribucion_dorm, {
    E47: 'DORMITORIO PROPIO \n(     )',
    F47: 'COMPARTE DORMITORIO \n(     )',
    G47: 'COMPARTE HABITACIÓN (     )'
  });

  const serv = Array.isArray(x.servicios) ? x.servicios : [];
  markCheck(ws, 'C48', serv.includes('agua'), 'AGUA POTABLE (     )');
  markCheck(ws, 'D48', serv.includes('desague'), 'DESAGÜE (     )');
  markCheck(ws, 'E48', serv.includes('luz'), 'ENERGÍA ELÉCTRICA (     )');
  markCheck(ws, 'F48', serv.includes('internet'), ' INTERNET (     )');

  markCheck(ws, 'D49', x.barreras === 'si', 'PRESENTA BARRERAS (     )');
  markCheck(ws, 'F49', x.barreras === 'no', 'NO PRESENTA BARRERAS (     )');
  // D50:H50 está combinada; el master es D50 (no E50)
  if (x.barreras_especificar) {
    const c = ws.getCell('D50');
    const base = cellText(c.value) || 'ESPECIFICAR:';
    c.value = /ESPECIFICAR/i.test(base)
      ? base.replace(/ESPECIFICAR:\s*/i, 'ESPECIFICAR: ') + x.barreras_especificar
      : 'ESPECIFICAR: ' + x.barreras_especificar;
    c.font = { name: 'Calibri', size: 12, bold: true, color: { argb: 'FF0000FF' } };
  }

  // —— Cierre ——
  setCell(ws, 'D51', x.diagnostico_social);
  setCell(ws, 'D52', x.derecho_mortuoria);
  setCell(ws, 'C54', fmtFecha(x.fecha_entrevista));
  setCell(ws, 'C55', x.profesional);
  setCell(ws, 'H54', x.firma);
}

/** Rellena la 1.ª fila de datos en «Registro Socios» (si existe la hoja). */
function rellenarRegistroSocios(ws) {
  const x = r();
  const row = 2; // fila 1 = encabezados
  const labels = {
    estado_civil: { soltero: 'Soltero(a)', casado: 'Casado(a)', viudo: 'Viudo(a)', conviviente: 'Conviviente', divorciado: 'Divorciado(a)', separado: 'Separado(a)' },
    tipo_familia: { unipersonal: 'Unipersonal', nuclear: 'Nuclear', monoparental: 'Monoparental', extensa: 'Extensa', otros: 'Otros' },
    tipo_discapacidad: { ceguera: 'Ceguera total', 'baja-vision': 'Baja visión' },
    condicion: { congenita: 'Congénita', adquirida: 'Adquirida' },
    seguro: { sis: 'SIS', essalud: 'EsSalud', privado: 'Privado', otros: 'Otros' },
    autonomia: { independiente: 'Independiente', parcial: 'Requiere apoyo parcial', permanente: 'Requiere apoyo permanente' },
    ingreso: {
      'sin-ingresos': 'Sin ingresos', 'menos-500': '< S/500', '500-900': 'S/500-900',
      '900-1500': 'S/900-1500', 'mas-1500': '> S/1500'
    },
    fuente_ingresos: {
      dependiente: 'Trabajo dependiente', independiente: 'Trabajo independiente',
      pension: 'Pensión/Jubilación', programa: 'Programa social',
      'apoyo-familiar': 'Apoyo familiar', otros: 'Otros'
    },
    sisfoh: {
      'no-tiene': 'No tiene', 'no-sabe': 'No sabe', pobre: 'Pobre',
      'pobre-extremo': 'Pobre extremo', 'no-pobre': 'No pobre'
    },
    tenencia: { propia: 'Propia', alquilada: 'Alquilada', cedida: 'Cedida', otro: 'Otro' },
    tipo_vivienda: {
      casa: 'Casa independiente', departamento: 'Departamento', cuarto: 'Cuarto',
      quinta: 'Quinta', multifamiliar: 'Multifamiliar', otro: 'Otro'
    },
    distribucion_dorm: {
      propio: 'Dormitorio propio', 'comparte-dorm': 'Comparte dormitorio',
      'comparte-hab': 'Comparte habitación'
    }
  };
  const L = (map, v) => (map && map[v]) || v || '';
  const serv = Array.isArray(x.servicios) ? x.servicios.join(', ') : (x.servicios || '');

  // Columnas alineadas a build_ficha_completa.py (A=N°, B=nombres, …)
  const vals = [
    1, // A N°
    x.nombres || '',
    x.dni || '',
    fmtFecha(x.fecha_nac) || '',
    x.edad != null && x.edad !== '' ? x.edad : '',
    x.sexo || '',
    x.depto_nac || '',
    x.prov_nac || '',
    x.dist_nac || '',
    L(labels.estado_civil, x.estado_civil),
    x.direccion || '',
    x.referencia || '',
    labelInstruccion(x.instruccion) || '',
    x.telefono || '',
    x.correo || '',
    x.c1_nombre || '',
    x.c1_par || '',
    x.c1_cel || '',
    x.c2_nombre || '',
    x.c2_par || '',
    x.c2_cel || '',
    siNoTxt(x.es_padre_madre),
    x.n_hijos != null && x.n_hijos !== '' ? x.n_hijos : '',
    x.n_hijos_disc != null && x.n_hijos_disc !== '' ? x.n_hijos_disc : '',
    x.otros_fam_disc || '',
    x.fam1_nombre || '', x.fam1_parentesco || '', x.fam1_edad || '', x.fam1_instruccion || '', x.fam1_ocupacion || '', siNoTxt(x.fam1_aporta), x.fam1_disc_enf || '',
    x.fam2_nombre || '', x.fam2_parentesco || '', x.fam2_edad || '', x.fam2_instruccion || '', x.fam2_ocupacion || '', siNoTxt(x.fam2_aporta), x.fam2_disc_enf || '',
    x.fam3_nombre || '', x.fam3_parentesco || '', x.fam3_edad || '', x.fam3_instruccion || '', x.fam3_ocupacion || '', siNoTxt(x.fam3_aporta), x.fam3_disc_enf || '',
    x.fam4_nombre || '', x.fam4_parentesco || '', x.fam4_edad || '', x.fam4_instruccion || '', x.fam4_ocupacion || '', siNoTxt(x.fam4_aporta), x.fam4_disc_enf || '',
    x.fam5_nombre || '', x.fam5_parentesco || '', x.fam5_edad || '', x.fam5_instruccion || '', x.fam5_ocupacion || '', siNoTxt(x.fam5_aporta), x.fam5_disc_enf || '',
    L(labels.tipo_familia, x.tipo_familia) + (x.tipo_familia_otro ? ' — ' + x.tipo_familia_otro : ''),
    L({ buena: 'Buena', regular: 'Regular', mala: 'Mala' }, x.relaciones_fam),
    L({ buena: 'Buena', regular: 'Regular', mala: 'Mala' }, x.comunicacion_fam),
    x.red_apoyo || '',
    x.apreciacion_prof || '',
    L(labels.tipo_discapacidad, x.tipo_discapacidad),
    L(labels.condicion, x.condicion),
    x.edad_adquisicion != null && x.edad_adquisicion !== '' ? x.edad_adquisicion : '',
    siNoTxt(x.certificado_disc),
    siNoTxt(x.carnet_conadis),
    siNoTxt(x.atencion_oftal),
    x.lugar_atencion || '',
    L(labels.seguro, x.seguro),
    x.seguro_otro || '',
    x.enfermedades_cronicas || '',
    L(labels.autonomia, x.autonomia),
    '', // detalle asistencia
    L(labels.ingreso, x.ingreso),
    siNoTxt(x.laborando),
    x.ocupacion || '',
    x.profesion || '',
    L(labels.fuente_ingresos, x.fuente_ingresos),
    x.programa_social_detalle || '',
    siNoTxt(x.otros_ingresos_hogar),
    x.principal_aportante || '',
    L(labels.sisfoh, x.sisfoh),
    L(labels.tenencia, x.tenencia) + (x.tenencia_otro ? ' — ' + x.tenencia_otro : ''),
    L(labels.tipo_vivienda, x.tipo_vivienda) + (x.tipo_vivienda_otro ? ' — ' + x.tipo_vivienda_otro : ''),
    x.n_ambientes != null && x.n_ambientes !== '' ? x.n_ambientes : '',
    L(labels.distribucion_dorm, x.distribucion_dorm),
    serv,
    siNoTxt(x.barreras),
    x.barreras_especificar || '',
    x.diagnostico_social || '',
    x.derecho_mortuoria || '',
    fmtFecha(x.fecha_entrevista) || '',
    x.profesional || '',
    x.firma || '',
    ''
  ];

  vals.forEach((v, i) => {
    if (v === undefined || v === null || v === '') return;
    const cell = ws.getCell(row, i + 1);
    cell.value = v;
    cell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FF0000FF' } };
  });
}

function nombreArchivoFicha(ext) {
  const nom = (respuestas.nombres || 'socio').toString()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^\w\s-]/g, '')
    .trim()
    .replace(/\s+/g, '_')
    .slice(0, 40) || 'socio';
  const dni = (respuestas.dni || '').toString().replace(/\D/g, '').slice(0, 12);
  const fecha = (respuestas.fecha_entrevista || new Date().toISOString().slice(0, 10)).replace(/-/g, '');
  return `Ficha_Social_UNCP_${nom}${dni ? '_' + dni : ''}_${fecha}.${ext}`;
}

function descargarBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => {
    URL.revokeObjectURL(url);
    a.remove();
  }, 1500);
}

// ═══════════════════════════════════════════
// BORRADOR
// ═══════════════════════════════════════════
function nuevaFicha() {
  if (!confirm('¿Borrar los datos actuales y empezar una ficha nueva?')) return;
  respuestas = {};
  cola = [];
  indice = 0;
  miembrosInsertados = 0;
  cancelarEscuchaProgramada();
  detenerMicrofono(false);
  detenerHabla();
  desactivarModoVozContinua(false);
  localStorage.removeItem('fichaUNCP_v3');
  localStorage.removeItem('fichaUNCP_v2');
  document.getElementById('pantalla-resumen').classList.remove('activa');
  document.getElementById('pantalla-entrevista').classList.remove('activa');
  document.getElementById('pantalla-inicio').classList.add('activa');
  // Volver a elegir modo en el inicio (no narrar aquí)
  setLecturaVoz(false, { persistir: false });
  mostrarMensaje('Ficha reiniciada. Elija Comenzar con voz o Comenzar sin voz.');
}

function guardarBorrador(silencioso) {
  const data = {
    respuestas,
    indice,
    miembrosInsertados,
    version: 3,
    lecturaVoz: !!lecturaVozActiva
  };
  localStorage.setItem('fichaUNCP_v3', JSON.stringify(data));
  // compat
  localStorage.setItem('fichaUNCP_v2', JSON.stringify(data));
  if (!silencioso) {
    if (lecturaVozActiva) hablar('Borrador guardado. Puede continuar más tarde.');
    else mostrarMensaje('Borrador guardado. Puede continuar más tarde.');
  }
}

async function cargarBorradorYContinuar() {
  let raw = localStorage.getItem('fichaUNCP_v3') || localStorage.getItem('fichaUNCP_v2');
  if (!raw) {
    mostrarMensaje('No hay borrador guardado.');
    return;
  }
  try {
    const data = JSON.parse(raw);
    respuestas = data.respuestas || {};
    indice = data.indice || 0;
    miembrosInsertados = data.miembrosInsertados || 0;
    cola = construirColaBase();
    const n = parseInt(respuestas.num_familia, 10) || 0;
    miembrosInsertados = n;
    if (indice >= cola.length) indice = Math.max(0, cola.length - 1);
    if (cola[indice] && !esPreguntaActiva(cola[indice])) {
      const adj = buscarIndiceActivo(indice, +1);
      indice = adj < cola.length ? adj : Math.max(0, buscarIndiceActivo(cola.length - 1, -1));
    }
    // Por defecto SIN narrador al continuar borrador (no se restaura voz sola)
    setLecturaVoz(false);

    document.getElementById('pantalla-inicio').classList.remove('activa');
    document.getElementById('pantalla-resumen').classList.remove('activa');
    document.getElementById('pantalla-entrevista').classList.add('activa');
    // Solo reactivar mic continuo si el modo es con voz
    if (lecturaVozActiva && soportaVoz() && contextoSeguro()) {
      await activarModoVozContinua({ silencioso: true, sinHablar: true });
    }
    actualizarBotonMicro();
    hablar(
      'Borrador recuperado. Continuamos desde donde se quedó.' +
      (lecturaVozActiva
        ? (modoVozContinua ? ' El micrófono escuchará solo después de cada pregunta.' : '')
        : ' Modo sin narrador: solo texto en pantalla.'),
      () => mostrarPreguntaActual()
    );
  } catch (e) {
    mostrarMensaje('El borrador está dañado. Inicie una ficha nueva.');
  }
}

// ═══════════════════════════════════════════
// GOOGLE SHEETS + PDF/EXCEL EN DRIVE (2026-08)
// Hoja:   https://docs.google.com/spreadsheets/d/1obx8kVIXXxk2P65LVXxHju3I1iUf3SsPgozlxKmZ5ek/edit
// Drive:  https://drive.google.com/drive/folders/1dqy2HuWf6IYVGasuyrtff2id7Jf3aRXv
// Guía:   CONFIGURAR-GOOGLE.md
// ═══════════════════════════════════════════
const SHEETS_SPREADSHEET_ID = '1obx8kVIXXxk2P65LVXxHju3I1iUf3SsPgozlxKmZ5ek';
const SHEETS_URL_KEY = 'fichaUNCP_sheetsUrl';
const SHEETS_HOJA_URL =
  'https://docs.google.com/spreadsheets/d/1obx8kVIXXxk2P65LVXxHju3I1iUf3SsPgozlxKmZ5ek/edit';
const SHEETS_DRIVE_FOLDER_URL =
  'https://drive.google.com/drive/folders/1dqy2HuWf6IYVGasuyrtff2id7Jf3aRXv';
/**
 * URL de la aplicación web (Apps Script) — /exec.
 * Hoja: 1obx8kVIXXxk2P65LVXxHju3I1iUf3SsPgozlxKmZ5ek
 * Drive: 1dqy2HuWf6IYVGasuyrtff2id7Jf3aRXv
 * También se puede sobrescribir en Admin → URL / config (localStorage).
 */
const SHEETS_WEBAPP_URL_DEFAULT =
  'https://script.google.com/macros/s/AKfycbw2Gg0qOVeBjiRl0N0uNlgxsdUkbiPV8D1wDGepXv7H3TbXtfECszqNGI95IN0y17Az/exec';

/** Columnas fijas + un bloque por cada familiar (1–7), en el orden de la ficha. */
function columnasGoogleSheets() {
  const base = [
    { id: '_timestamp', header: 'Fecha registro' },
    { id: 'fecha_entrevista', header: 'Fecha entrevista' },
    { id: 'profesional', header: 'Profesional' },
    { id: 'nombres', header: 'Apellidos y nombres' },
    { id: 'dni', header: 'DNI' },
    { id: 'fecha_nac', header: 'Fecha nacimiento' },
    { id: 'edad', header: 'Edad' },
    { id: 'sexo', header: 'Sexo' },
    { id: 'depto_nac', header: 'Depto nacimiento' },
    { id: 'prov_nac', header: 'Provincia nacimiento' },
    { id: 'dist_nac', header: 'Distrito nacimiento' },
    { id: 'estado_civil', header: 'Estado civil' },
    { id: 'direccion', header: 'Dirección' },
    { id: 'referencia', header: 'Referencia' },
    { id: 'instruccion', header: 'Nivel educativo' },
    { id: 'telefono', header: 'Teléfono' },
    { id: 'correo', header: 'Correo' },
    { id: 'c1_nombre', header: 'Contacto 1 nombre' },
    { id: 'c1_par', header: 'Contacto 1 parentesco' },
    { id: 'c1_cel', header: 'Contacto 1 celular' },
    { id: 'c2_nombre', header: 'Contacto 2 nombre' },
    { id: 'c2_par', header: 'Contacto 2 parentesco' },
    { id: 'c2_cel', header: 'Contacto 2 celular' },
    { id: 'es_padre_madre', header: 'Es padre/madre' },
    { id: 'n_hijos', header: 'N° hijos' },
    { id: 'n_hijos_disc', header: 'N° hijos disc. visual' },
    { id: 'otros_fam_disc', header: 'Otros fam. disc. visual' },
    { id: 'num_familia', header: 'N° convivientes' },
    { id: 'tipo_familia', header: 'Tipo de familia' },
    { id: 'tipo_familia_otro', header: 'Tipo familia (otro)' },
    { id: 'relaciones_fam', header: 'Relaciones familiares' },
    { id: 'comunicacion_fam', header: 'Comunicación familiar' },
    { id: 'red_apoyo', header: 'Red de apoyo' },
    { id: 'apreciacion_prof', header: 'Apreciación profesional' },
    { id: 'tipo_discapacidad', header: 'Tipo discapacidad' },
    { id: 'condicion', header: 'Origen discapacidad' },
    { id: 'edad_adquisicion', header: 'Edad adquisición' },
    { id: 'certificado_disc', header: 'Certificado discapacidad' },
    { id: 'carnet_conadis', header: 'Carnet CONADIS' },
    { id: 'atencion_oftal', header: 'Atención oftalmológica' },
    { id: 'lugar_atencion', header: 'Lugar atención' },
    { id: 'seguro', header: 'Seguro de salud' },
    { id: 'seguro_otro', header: 'Seguro (otro)' },
    { id: 'enfermedades_cronicas', header: 'Enfermedades crónicas' },
    { id: 'autonomia', header: 'Autonomía' },
    { id: 'ingreso', header: 'Ingreso mensual' },
    { id: 'laborando', header: 'Laborando' },
    { id: 'ocupacion', header: 'Ocupación' },
    { id: 'profesion', header: 'Profesión' },
    { id: 'fuente_ingresos', header: 'Fuente de ingresos' },
    { id: 'programa_social_detalle', header: 'Programa social detalle' },
    { id: 'otros_ingresos_hogar', header: 'Otros ingresos hogar' },
    { id: 'principal_aportante', header: 'Principal aportante' },
    { id: 'sisfoh', header: 'SISFOH' },
    { id: 'tenencia', header: 'Tenencia vivienda' },
    { id: 'tenencia_otro', header: 'Tenencia (otro)' },
    { id: 'tipo_vivienda', header: 'Tipo vivienda' },
    { id: 'tipo_vivienda_otro', header: 'Tipo vivienda (otro)' },
    { id: 'n_ambientes', header: 'N° ambientes' },
    { id: 'distribucion_dorm', header: 'Distribución dormitorio' },
    { id: 'servicios', header: 'Servicios básicos' },
    { id: 'barreras', header: 'Barreras arquitectónicas' },
    { id: 'barreras_especificar', header: 'Barreras (especificar)' },
    { id: 'diagnostico_social', header: 'Diagnóstico social' },
    { id: 'derecho_mortuoria', header: 'Derecho mortuoria' },
    { id: 'firma', header: 'Firma / huella' }
  ];

  const famLabels = {
    nombre: 'nombres',
    parentesco: 'parentesco',
    edad: 'edad',
    instruccion: 'nivel educativo',
    ocupacion: 'ocupación',
    aporta: 'aporta ingresos',
    disc_enf: 'disc./enfermedad'
  };

  for (let n = 1; n <= MAX_MIEMBROS_FAMILIA; n++) {
    CAMPOS_MIEMBRO.forEach(campo => {
      base.push({
        id: `fam${n}_${campo}`,
        header: `Familiar ${n} — ${famLabels[campo] || campo}`
      });
    });
  }
  return base;
}

function valorParaSheets(id, raw) {
  if (id === '_timestamp') {
    return new Date().toISOString();
  }
  if (raw === undefined || raw === null) return '';
  if (Array.isArray(raw)) return raw.join(', ');

  // Etiquetas legibles desde la cola (incluye miembros) o PREGUNTAS base
  const lista = cola.length ? cola : construirColaBase();
  const q = lista.find(p => p.id === id) || PREGUNTAS.find(p => p.id === id);
  if (q && q.opciones) {
    return etiquetaDeValor(q, raw);
  }
  if (String(id).endsWith('_aporta')) {
    if (raw === 'si') return 'Sí';
    if (raw === 'no') return 'No';
  }
  if (id === 'instruccion') return labelInstruccion(raw) || String(raw);
  if (id === 'fecha_nac' || id === 'fecha_entrevista') return fmtFecha(raw) || String(raw);
  return String(raw);
}

function construirPayloadSheets() {
  const cols = columnasGoogleSheets();
  const x = respuestas || {};
  const headers = cols.map(c => c.header);
  const values = cols.map(c => valorParaSheets(c.id, c.id === '_timestamp' ? null : x[c.id]));
  return {
    action: 'append',
    spreadsheetId: SHEETS_SPREADSHEET_ID,
    headers,
    values,
    pdfName: nombreArchivoPdfFicha(),
    meta: {
      nombres: x.nombres || '',
      dni: x.dni || '',
      fecha_entrevista: x.fecha_entrevista || ''
    }
  };
}

/** Nombre del Word en Drive (mismo base que el PDF). */
function nombreArchivoWordFicha() {
  return nombreArchivoPdfFicha().replace(/\.pdf$/i, '.doc');
}

function blobToBase64(blob) {
  return new Promise((resolve, reject) => {
    const fr = new FileReader();
    fr.onload = () => {
      const s = String(fr.result || '');
      const i = s.indexOf(',');
      resolve(i >= 0 ? s.slice(i + 1) : s);
    };
    fr.onerror = reject;
    fr.readAsDataURL(blob);
  });
}

/**
 * Genera Word (.doc HTML) con la ficha en cuadros y todos los datos.
 * Texto real, se abre en Word / Google Docs.
 * @returns {Promise<string|null>} base64
 */
async function generarWordFichaBase64() {
  const x = respuestas || {};
  const logoUrl = await obtenerLogoDataUrl();
  const escH = (s) => String(s == null ? '' : s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');

  const L = (t) => `<td style="border:1px solid #000;padding:4px 6px;font-weight:bold;background:#f0f0f0;width:22%;">${escH(t)}</td>`;
  const V = (t, span) => `<td style="border:1px solid #000;padding:4px 6px;"${span ? ` colspan="${span}"` : ''}>${escH(t)}</td>`;
  const row = (cells) => `<tr>${cells}</tr>`;
  const sec = (t) => `<p style="margin:12px 0 4px;font-weight:bold;color:#666;font-size:10pt;">${escH(t)}</p>`;

  const nFam = parseInt(x.num_familia, 10) || 0;
  const filasFam = Math.max(nFam, 3);
  let famRows = '';
  for (let i = 1; i <= filasFam; i++) {
    famRows += row([
      V(x[`fam${i}_nombre`]),
      V(x[`fam${i}_parentesco`]),
      V(x[`fam${i}_edad`]),
      V(x[`fam${i}_instruccion`]),
      V(x[`fam${i}_ocupacion`]),
      V(x[`fam${i}_aporta`] === 'si' ? 'Sí' : x[`fam${i}_aporta`] === 'no' ? 'No' : ''),
      V(x[`fam${i}_disc_enf`])
    ].join(''));
  }
  const serv = Array.isArray(x.servicios) ? x.servicios : [];

  const html = `<!DOCTYPE html>
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word">
<head>
<meta charset="utf-8">
<title>Ficha Social UNCP</title>
<!--[if gte mso 9]><xml><w:WordDocument><w:View>Print</w:View></w:WordDocument></xml><![endif]-->
<style>
  body { font-family: Arial, Helvetica, sans-serif; font-size: 10pt; color: #000; }
  table { border-collapse: collapse; width: 100%; margin-bottom: 4px; }
  td, th { border: 1px solid #000; padding: 4px 6px; vertical-align: middle; }
  th { background: #eee; font-size: 8pt; }
  .hdr { text-align: center; border: 2px solid #000; padding: 10px; margin-bottom: 10px; }
  .hdr h1 { margin: 6px 0 2px; font-size: 14pt; }
</style>
</head>
<body>
<div class="hdr">
  ${logoUrl ? `<div style="background:#000;display:inline-block;padding:6px 12px;"><img src="${logoUrl}" height="40" alt="UNCP"></div>` : ''}
  <h1>FICHA SOCIAL UNCP</h1>
  <div>Unión Nacional de Ciegos del Perú — Formato oficial</div>
</div>

${sec('I. DATOS GENERALES DEL SOCIO(A)')}
<table>
${row(L('APELLIDOS Y NOMBRES') + V(x.nombres, 5))}
${row(L('N° DNI / CARNÉ') + V(x.dni) + L('FECHA NAC.') + V(fmtFecha(x.fecha_nac)) + L('EDAD') + V(x.edad))}
${row(L('SEXO') + V(`M (${pdfMarca(x.sexo, 'M')})   F (${pdfMarca(x.sexo, 'F')})`, 5))}
${row(L('LUGAR NACIMIENTO') + V('DEP: ' + pdfTxt(x.depto_nac)) + V('PROV: ' + pdfTxt(x.prov_nac), 2) + V('DIST: ' + pdfTxt(x.dist_nac), 2))}
${row(L('ESTADO CIVIL') + V([
  pdfOpc('SOLTERO/A', x.estado_civil, 'soltero'),
  pdfOpc('CASADO/A', x.estado_civil, 'casado'),
  pdfOpc('VIUDO/A', x.estado_civil, 'viudo'),
  pdfOpc('CONVIVIENTE', x.estado_civil, 'conviviente'),
  pdfOpc('DIVORCIADO/A', x.estado_civil, 'divorciado'),
  pdfOpc('SEPARADO/A', x.estado_civil, 'separado')
].join('  '), 5))}
${row(L('DIRECCIÓN ACTUAL') + V(x.direccion, 5))}
${row(L('REFERENCIA') + V(x.referencia, 3) + L('NIVEL EDUCATIVO') + V(labelInstruccion(x.instruccion)))}
${row(L('CELULAR') + V(x.telefono, 2) + L('CORREO') + V(x.correo, 2))}
</table>

${sec('CONTACTO DE EMERGENCIA')}
<table>
${row(L('CONTACTO 1 — NOMBRE') + V(x.c1_nombre) + L('PARENTESCO') + V(x.c1_par) + L('CELULAR') + V(x.c1_cel))}
${row(L('CONTACTO 2 — NOMBRE') + V(x.c2_nombre) + L('PARENTESCO') + V(x.c2_par) + L('CELULAR') + V(x.c2_cel))}
</table>

${sec('II. CARACTERIZACIÓN FAMILIAR')}
<table>
${row(L('¿ES USTED PADRE O MADRE?') + V(pdfSiNo(x.es_padre_madre)) + L('N° TOTAL HIJOS') + V(x.n_hijos) + L('N° HIJOS DISC. VISUAL') + V(x.n_hijos_disc))}
${row(L('OTROS FAM. DISC. VISUAL') + V(x.otros_fam_disc, 5))}
</table>

${sec('III. COMPOSICIÓN DEL NÚCLEO FAMILIAR')}
<table>
<tr>
  <th>APELLIDOS Y NOMBRES</th><th>PARENTESCO</th><th>EDAD</th><th>NIVEL EDUCATIVO</th>
  <th>OCUPACIÓN</th><th>¿APORTA?</th><th>DISC./ENFERMEDAD</th>
</tr>
${famRows}
</table>
<table>
${row(L('TIPO DE FAMILIA') + V([
  pdfOpc('UNIPERSONAL', x.tipo_familia, 'unipersonal'),
  pdfOpc('NUCLEAR', x.tipo_familia, 'nuclear'),
  pdfOpc('MONOPARENTAL', x.tipo_familia, 'monoparental'),
  pdfOpc('EXTENSA', x.tipo_familia, 'extensa'),
  'OTROS: ' + pdfTxt(x.tipo_familia_otro)
].join('  '), 5))}
</table>

${sec('IV. APRECIACIÓN DEL NÚCLEO FAMILIAR')}
<table>
${row(L('RELACIONES FAMILIARES') + V([pdfOpc('BUENA', x.relaciones_fam, 'buena'), pdfOpc('REGULAR', x.relaciones_fam, 'regular'), pdfOpc('MALA', x.relaciones_fam, 'mala')].join('  '), 5))}
${row(L('COMUNICACIÓN FAMILIAR') + V([pdfOpc('BUENA', x.comunicacion_fam, 'buena'), pdfOpc('REGULAR', x.comunicacion_fam, 'regular'), pdfOpc('MALA', x.comunicacion_fam, 'mala')].join('  '), 5))}
${row(L('PRINCIPAL RED DE APOYO') + V(x.red_apoyo, 5))}
${row(L('APRECIACIÓN PROFESIONAL') + V(x.apreciacion_prof, 5))}
</table>

${sec('V. CONDICIÓN DE DISCAPACIDAD VISUAL')}
<table>
${row(L('CONDICIÓN') + V([pdfOpc('CEGUERA TOTAL', x.tipo_discapacidad, 'ceguera'), pdfOpc('BAJA VISIÓN', x.tipo_discapacidad, 'baja-vision')].join('  '), 2) + L('ORIGEN') + V([pdfOpc('CONGÉNITA', x.condicion, 'congenita'), pdfOpc('ADQUIRIDA', x.condicion, 'adquirida')].join('  '), 2))}
${row(L('EDAD ADQUISICIÓN') + V(x.edad_adquisicion) + L('CERT. DISCAPACIDAD') + V(x.certificado_disc === 'si' ? 'Sí' : x.certificado_disc === 'no' ? 'No' : '') + L('CARNET CONADIS') + V(x.carnet_conadis === 'si' ? 'Sí' : x.carnet_conadis === 'no' ? 'No' : ''))}
${row(L('ATENCIÓN OFTALMOLÓGICA') + V(pdfSiNo(x.atencion_oftal), 2) + L('LUGAR / INSTITUCIÓN') + V(x.lugar_atencion, 2))}
</table>

${sec('VI. ESTADO DE SALUD')}
<table>
${row(L('SEGURO DE SALUD') + V([pdfOpc('SIS', x.seguro, 'sis'), pdfOpc('ESSALUD', x.seguro, 'essalud'), pdfOpc('PRIVADO', x.seguro, 'privado'), pdfOpc('OTROS', x.seguro, 'otros')].join('  ') + (x.seguro_otro ? ' ' + x.seguro_otro : ''), 5))}
${row(L('ENFERMEDADES CRÓNICAS') + V(x.enfermedades_cronicas, 5))}
${row(L('AUTONOMÍA') + V([pdfOpc('INDEPENDIENTE', x.autonomia, 'independiente'), pdfOpc('APOYO PARCIAL', x.autonomia, 'parcial'), pdfOpc('APOYO PERMANENTE', x.autonomia, 'permanente')].join('  '), 5))}
</table>

${sec('VII. SITUACIÓN LABORAL Y ECONÓMICA')}
<table>
${row(L('INGRESO MENSUAL') + V([pdfOpc('SIN INGRESOS', x.ingreso, 'sin-ingresos'), pdfOpc('< S/500', x.ingreso, 'menos-500'), pdfOpc('S/500-900', x.ingreso, '500-900'), pdfOpc('S/900-1500', x.ingreso, '900-1500'), pdfOpc('> S/1500', x.ingreso, 'mas-1500')].join('  '), 5))}
${row(L('¿LABORA?') + V(pdfSiNo(x.laborando)) + L('OCUPACIÓN') + V(x.ocupacion) + L('PROFESIÓN') + V(x.profesion))}
${row(L('FUENTE INGRESOS') + V([pdfOpc('TRAB. DEPENDIENTE', x.fuente_ingresos, 'dependiente'), pdfOpc('TRAB. INDEPENDIENTE', x.fuente_ingresos, 'independiente'), pdfOpc('PENSIÓN', x.fuente_ingresos, 'pension'), pdfOpc('PROGRAMA SOCIAL', x.fuente_ingresos, 'programa'), pdfOpc('APOYO FAMILIAR', x.fuente_ingresos, 'apoyo-familiar'), pdfOpc('OTROS', x.fuente_ingresos, 'otros')].join('  ') + (x.programa_social_detalle ? ' — ' + x.programa_social_detalle : ''), 5))}
${row(L('¿OTROS INGRESOS HOGAR?') + V(pdfSiNo(x.otros_ingresos_hogar), 2) + L('PRINCIPAL APORTANTE') + V(x.principal_aportante, 2))}
${row(L('SISFOH') + V([pdfOpc('NO TIENE', x.sisfoh, 'no-tiene'), pdfOpc('NO SABE', x.sisfoh, 'no-sabe'), pdfOpc('POBRE', x.sisfoh, 'pobre'), pdfOpc('POBRE EXTREMO', x.sisfoh, 'pobre-extremo'), pdfOpc('NO POBRE', x.sisfoh, 'no-pobre')].join('  '), 5))}
</table>

${sec('VIII. CONDICIÓN DE VIVIENDA')}
<table>
${row(L('TENENCIA') + V([pdfOpc('PROPIA', x.tenencia, 'propia'), pdfOpc('ALQUILADA', x.tenencia, 'alquilada'), pdfOpc('CEDIDA', x.tenencia, 'cedida'), 'OTRO: ' + pdfTxt(x.tenencia_otro)].join('  '), 5))}
${row(L('TIPO VIVIENDA') + V([pdfOpc('CASA INDEPENDIENTE', x.tipo_vivienda, 'casa'), pdfOpc('DEPARTAMENTO', x.tipo_vivienda, 'departamento'), pdfOpc('CUARTO', x.tipo_vivienda, 'cuarto'), pdfOpc('QUINTA', x.tipo_vivienda, 'quinta'), pdfOpc('MULTIFAMILIAR', x.tipo_vivienda, 'multifamiliar'), 'OTRO: ' + pdfTxt(x.tipo_vivienda_otro)].join('  '), 5))}
${row(L('N° AMBIENTES') + V(x.n_ambientes) + L('DISTRIBUCIÓN') + V([pdfOpc('DORM. PROPIO', x.distribucion_dorm, 'propio'), pdfOpc('COMPARTE DORM.', x.distribucion_dorm, 'comparte-dorm'), pdfOpc('COMPARTE HAB.', x.distribucion_dorm, 'comparte-hab')].join('  '), 3))}
${row(L('SERVICIOS BÁSICOS') + V([pdfOpc('AGUA', serv.includes('agua') ? 'agua' : '', 'agua'), pdfOpc('DESAGÜE', serv.includes('desague') ? 'desague' : '', 'desague'), pdfOpc('LUZ', serv.includes('luz') ? 'luz' : '', 'luz'), pdfOpc('INTERNET', serv.includes('internet') ? 'internet' : '', 'internet')].join('  '), 5))}
${row(L('ACCESIBILIDAD') + V([pdfOpc('PRESENTA BARRERAS', x.barreras, 'si'), pdfOpc('NO PRESENTA', x.barreras, 'no')].join('  '), 2) + L('ESPECIFICAR') + V(x.barreras_especificar, 2))}
${row(L('DIAGNÓSTICO SOCIAL') + V(x.diagnostico_social, 5))}
${row(L('DERECHO MORTUORIA') + V(x.derecho_mortuoria, 5))}
</table>

${sec('REGISTRO ADMINISTRATIVO')}
<table>
${row(L('FECHA ENTREVISTA') + V(fmtFecha(x.fecha_entrevista), 2) + L('PROFESIONAL') + V(x.profesional, 2))}
${row(L('FIRMA O HUELLA') + V(x.firma || '', 5))}
</table>
</body></html>`;

  try {
    // BOM UTF-8 para que Word abra tildes bien
    const bom = '\ufeff';
    const blob = new Blob([bom + html], { type: 'application/msword' });
    const b64 = await blobToBase64(blob);
    console.log('Word ficha OK, base64 length=', b64.length);
    return b64;
  } catch (e) {
    console.error('Word ficha', e);
    return null;
  }
}

/**
 * Payload: datos + PDF + Excel (misma ficha) + Word para Drive.
 */
async function construirPayloadSheetsConPdf() {
  rellenarPlantillaPDF();
  const payload = construirPayloadSheets();
  payload.wordName = nombreArchivoWordFicha();
  payload.excelName = nombreArchivoFicha('xlsx');
  payload.pdfName = nombreArchivoPdfFicha();

  setEstadoSheets('📄 Generando PDF (plantilla impresa) y Excel…', false);

  try {
    // Misma plantilla que "Imprimir ficha" (#plantilla-pdf)
    let b64 = await generarPdfVisualBase64();
    if (!b64 || b64.length < 500) {
      await new Promise(r => setTimeout(r, 400));
      b64 = await generarPdfVisualBase64();
    }
    if (b64 && b64.length > 500) {
      payload.pdfBase64 = b64;
      payload.pdfOrigen = 'ficha_impresa';
      payload.generarPdfTexto = false;
    } else {
      payload.generarPdfTexto = true;
      payload.pdfOrigen = 'datos_texto';
    }
  } catch (e) {
    console.warn(e);
    payload.generarPdfTexto = true;
    payload.pdfOrigen = 'datos_texto';
  }

  // Excel = misma versión de datos de la ficha (plantilla oficial)
  try {
    setEstadoSheets('📊 Generando Excel (misma ficha que el PDF)…', false);
    const excelOut = await generarExcelBase64();
    if (excelOut && excelOut.base64 && excelOut.base64.length > 200) {
      payload.excelBase64 = excelOut.base64;
      payload.excelName = excelOut.nombre || nombreArchivoFicha('xlsx');
    }
  } catch (eX) {
    console.warn('Excel no generado', eX);
  }

  try {
    const w64 = await generarWordFichaBase64();
    if (w64 && w64.length > 100) {
      payload.wordBase64 = w64;
      payload.wordName = nombreArchivoWordFicha();
    }
  } catch (eW) {
    console.warn('Word no generado', eW);
  }

  if (payload.pdfBase64 || payload.excelBase64 || payload.wordBase64) {
    setEstadoSheets('✓ Archivos listos. Subiendo a Drive (PDF + Excel + Word)…', false);
  } else {
    setEstadoSheets('⚠ No se generaron archivos de ficha; se intentará guardar solo la fila.', true);
  }
  return payload;
}

function obtenerUrlSheets() {
  // Prioridad: URL fija en el código (nueva implementación).
  // localStorage solo si no hay default (evita quedarse con /exec viejo).
  const def = (SHEETS_WEBAPP_URL_DEFAULT || '').trim();
  if (def) return def;
  try {
    return (localStorage.getItem(SHEETS_URL_KEY) || '').trim();
  } catch (e) {
    return '';
  }
}

function guardarUrlSheets() {
  const input = document.getElementById('input-sheets-url');
  const url = (input && input.value ? input.value : '').trim();
  if (!url) {
    setEstadoSheets('⚠ Pegue la URL de la aplicación web de Apps Script.', true);
    hablar('Pegue la URL de la aplicación web de Apps Script.');
    return;
  }
  if (!/^https:\/\/script\.google\.com\//i.test(url) && !/^https:\/\/.+\/macros\/s\//i.test(url)) {
    setEstadoSheets('⚠ La URL no parece de Apps Script (script.google.com).', true);
    hablar('La URL no parece de Apps Script. Verifique e intente de nuevo.');
    return;
  }
  try {
    localStorage.setItem(SHEETS_URL_KEY, url);
  } catch (e) {
    setEstadoSheets('No se pudo guardar la URL en el navegador.', true);
    return;
  }
  setEstadoSheets('✓ URL guardada. Ya puede enviar fichas a Google Sheets.', false);
  hablar('URL de Google Sheets guardada.');
  actualizarEstadoSheetsUI();
  toggleConfigSheets(false);
}

function toggleConfigSheets(forzar) {
  if (!esModoAdmin()) {
    pedirClaveAdmin();
    if (!esModoAdmin()) return;
  }
  const box = document.getElementById('config-sheets-box');
  if (!box) return;
  const abrir = forzar === true || (forzar !== false && box.style.display === 'none');
  box.style.display = abrir ? 'block' : 'none';
  if (abrir) {
    const input = document.getElementById('input-sheets-url');
    if (input) {
      input.value = obtenerUrlSheets();
      try { input.focus(); } catch (e) { /* ignore */ }
    }
  }
}

function setEstadoSheets(texto, esError) {
  // Usuario normal: mensajes simples, sin detalles técnicos
  let textoUsuario = texto;
  if (!esModoAdmin()) {
    if (esError) {
      if (/script|v4|URL|Apps Script|columnas|implement/i.test(texto)) {
        textoUsuario = 'No se pudo completar el guardado. Avise al personal de la UNCP.';
      }
    } else if (/Generando|Subiendo|Preparando|PDF listo|Archivos listos/i.test(texto)) {
      textoUsuario = texto
        .replace(/Google Sheets/gi, 'la nube')
        .replace(/Drive/gi, 'la nube');
    } else if (/guardad|listo|OK|✓/i.test(texto)) {
      // dejar mensaje de éxito legible
    }
  }
  const el = document.getElementById('estado-sheets');
  if (el) {
    el.textContent = textoUsuario;
    el.style.color = esError ? 'var(--error)' : 'var(--text-soft)';
  }
  const msg = document.getElementById('mensaje-resumen');
  if (msg && textoUsuario) {
    msg.textContent = (esError ? '⚠ ' : '✅ ') + String(textoUsuario).replace(/^[⚠✓✅]\s*/, '');
  }
}

function actualizarEstadoSheetsUI() {
  actualizarUIAdmin();
  const url = obtenerUrlSheets();
  const input = document.getElementById('input-sheets-url');
  if (input && !input.value) input.value = url;
  if (esModoAdmin()) {
    if (url) {
      setEstadoSheets('✓ URL configurada. Guardado automático al terminar.', false);
    } else {
      setEstadoSheets('Admin: configure la URL /exec del Apps Script.', true);
    }
  } else {
    setEstadoSheets(
      url
        ? 'Al terminar se guarda la ficha en la nube (hoja, PDF, Excel y Word).'
        : 'Listo para guardar la ficha.',
      false
    );
  }
}

function mostrarAvisoScriptViejo(mostrar) {
  // Solo el administrador ve diagnósticos técnicos
  const el = document.getElementById('aviso-script-viejo');
  if (!el) return;
  if (!esModoAdmin()) {
    el.style.display = 'none';
    return;
  }
  if (mostrar) {
    el.innerHTML =
      '⚠ <strong>Diagnóstico admin:</strong> el script de Google puede estar desactualizado o la URL es incorrecta. ' +
      'Use Diagnosticar y revise la URL /exec. Versión esperada: v8 o v9.';
    el.style.display = 'block';
  } else {
    el.style.display = 'none';
  }
}

/**
 * Diagnóstico (solo admin): conexión, versión del script, hoja y carpeta.
 */
async function diagnosticarPdfDrive() {
  if (!esModoAdmin()) {
    pedirClaveAdmin();
    if (!esModoAdmin()) return;
  }
  const url = obtenerUrlSheets();
  if (!url) {
    setEstadoSheets('⚠ No hay URL de Apps Script configurada.', true);
    hablar('No hay URL del script de Google configurada.');
    return;
  }
  setEstadoSheets('🔎 Diagnosticando script de Google…', false);
  hablar('Diagnosticando por qué no se guarda el PDF.');
  try {
    const resp = await fetch(url, { method: 'GET', redirect: 'follow', cache: 'no-store' });
    const text = await resp.text();
    let data = null;
    try { data = JSON.parse(text); } catch (e) { /* ignore */ }

    console.log('Diagnóstico Apps Script GET:', text);

    if (!resp.ok) {
      setEstadoSheets('Error HTTP ' + resp.status + ' al consultar el script.', true);
      mostrarAvisoScriptViejo(true);
      hablar('Error al consultar el script de Google.');
      return;
    }

    // Versión NUEVA
    if (data && data.pdfReady && data.version) {
      mostrarAvisoScriptViejo(false);
      if (data.folderAccessible === false) {
        setEstadoSheets(
          'Script NUEVO OK, pero sin acceso a la carpeta Drive. Comparta la carpeta como EDITOR con la cuenta del script. ' +
          (data.folderError || ''),
          true
        );
        hablar('El script está actualizado, pero no tiene permiso de editor en la carpeta de Drive. Comparta la carpeta con la cuenta del script.');
        return;
      }
      setEstadoSheets(
        '✓ Script listo para PDF. Versión ' + data.version +
        (data.folderAccessible ? '. Carpeta Drive accesible.' : '') +
        ' Ya puede Guardar ficha + PDF.',
        false
      );
      hablar('Diagnóstico bien. El script ya puede guardar el PDF en Drive. Pulse Guardar ficha y PDF.');
      return;
    }

    // Versión VIEJA (lo que está pasando ahora)
    mostrarAvisoScriptViejo(true);
    setEstadoSheets(
      '⚠ CONFIRMADO: script VIEJO en Google. Respuesta: ' +
      (data && data.message ? data.message : text.slice(0, 120)) +
      ' → Solo guarda filas. NO crea PDF. Ejecute «Actualizar script Google.bat», pegue código, Implementar → Nueva versión.',
      true
    );
    hablar(
      'Confirmado: el script en Google está desactualizado. Por eso no se guarda el PDF. ' +
      'En su computadora ejecute el archivo Actualizar script Google punto bat. ' +
      'Luego en la página de Apps Script: control V para pegar, control S para guardar, ' +
      'Implementar, Administrar implementaciones, lápiz, Nueva versión, Implementar. ' +
      'Después pulse otra vez Diagnosticar PDF.'
    );
  } catch (err) {
    console.error(err);
    setEstadoSheets('No se pudo diagnosticar: ' + (err && err.message ? err.message : err), true);
    hablar('No se pudo conectar con Google para diagnosticar.');
  }
}

async function probarConexionSheets() {
  return diagnosticarPdfDrive();
}

/**
 * Interpreta respuesta de Apps Script (JSON, HTML de login, 401, etc.).
 */
function interpretarRespuestaAppsScript(resp, text) {
  const t = String(text || '');
  let data = null;
  try { data = JSON.parse(t); } catch (e) { data = null; }

  const looksLogin =
    /accounts\.google\.com|Sign in|Iniciar sesi[oó]n|ServiceLogin/i.test(t) ||
    (t.trim().startsWith('<!doctype html') && !data);

  let authError = false;
  if (resp && (resp.status === 401 || resp.status === 403)) authError = true;
  if (looksLogin) authError = true;

  return {
    data,
    text: t,
    authError,
    looksLogin,
    okHttp: !!(resp && resp.ok),
    status: resp ? resp.status : 0
  };
}

/** Mensaje claro cuando Google pide login / la app no es pública */
function mensajeErrorAuthAppsScript() {
  return (
    'Google rechazó el acceso (401). En Apps Script de la hoja: ' +
    'Implementar → Administrar implementaciones → lápiz → ' +
    '«Quién tiene acceso» = Cualquier persona → Nueva versión → Implementar. ' +
    'Luego vuelva a pegar la URL /exec y pulse Diagnosticar.'
  );
}

/**
 * Envía la ficha actual a la hoja de Google.
 * @param {{ silencioso?: boolean, auto?: boolean }} opts
 */
async function enviarAGoogleSheets(opts) {
  opts = opts || {};
  const url = obtenerUrlSheets();
  if (!url) {
    if (!opts.silencioso) {
      setEstadoSheets('⚠ Configure primero la URL de Apps Script (botón Configurar).', true);
      hablar('Configure primero Google Sheets con la URL de la aplicación web.');
      toggleConfigSheets(true);
    }
    return { ok: false, error: 'sin_url' };
  }

  const btn = document.getElementById('btn-guardar-sheets');
  if (btn) btn.disabled = true;
  setEstadoSheets('☁ Conectando con Google…', false);
  if (!opts.auto && !opts.silencioso) {
    hablar('Guardando el PDF y el Excel en Google Drive, y la ficha en la hoja.');
  }

  /**
   * POST a Apps Script (text/plain evita preflight CORS).
   * No usa mode:cors explícito; sigue redirecciones.
   */
  async function postJson(body, timeoutMs) {
    timeoutMs = timeoutMs || 120000;
    const ctrl = typeof AbortController !== 'undefined' ? new AbortController() : null;
    const timer = ctrl ? setTimeout(function () { try { ctrl.abort(); } catch (e) { /* */ } }, timeoutMs) : null;
    try {
      const resp = await fetch(url, {
        method: 'POST',
        redirect: 'follow',
        credentials: 'omit',
        cache: 'no-store',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify(body),
        signal: ctrl ? ctrl.signal : undefined
      });
      const text = await resp.text();
      const parsed = interpretarRespuestaAppsScript(resp, text);
      return {
        resp,
        text: parsed.text,
        data: parsed.data,
        authError: parsed.authError,
        okHttp: parsed.okHttp,
        status: parsed.status
      };
    } finally {
      if (timer) clearTimeout(timer);
    }
  }

  // —— 0) Ping: si falla aquí, no es “internet” sino acceso al script ——
  try {
    setEstadoSheets('☁ Comprobando acceso al script de Google…', false);
    const rPing = await postJson({ action: 'ping' }, 25000);
    if (rPing.authError || (rPing.text && /accounts\.google/i.test(rPing.text) && !rPing.data)) {
      setEstadoSheets(mensajeErrorAuthAppsScript(), true);
      hablar('Google pide permiso. La aplicación web debe ser de acceso para cualquier persona.');
      if (btn) btn.disabled = false;
      return { ok: false, error: 'auth_401' };
    }
    if (!rPing.data || !rPing.data.ok) {
      // Algunas implementaciones no tienen ping: seguimos
      console.warn('Ping no OK, se continúa', rPing.status, (rPing.text || '').slice(0, 120));
    }
  } catch (ePing) {
    console.error('Ping falló', ePing);
    const msg =
      'No se pudo conectar con Google (Failed to fetch). ' +
      'Causas habituales: 1) Apps Script no es «Cualquier persona», ' +
      '2) URL /exec incorrecta o implementación vieja, ' +
      '3) sin internet o bloqueo de red. ' +
      'Abra la URL /exec en el navegador: debe verse JSON, no pantalla de login.';
    setEstadoSheets(msg, true);
    hablar('No se pudo conectar con Google. Revise que la aplicación web sea pública.');
    if (btn) btn.disabled = false;
    return { ok: false, error: 'network_ping', detail: String(ePing && ePing.message ? ePing.message : ePing) };
  }

  let payload;
  try {
    payload = await construirPayloadSheetsConPdf();
  } catch (ePrep) {
    console.warn(ePrep);
    payload = construirPayloadSheets();
    payload.generarPdfTexto = true;
  }

  // Limitar tamaño: Google/fetch fallan con cuerpos enormes
  const MAX_B64 = 8 * 1024 * 1024; // ~8 MB base64
  if (payload.pdfBase64 && payload.pdfBase64.length > MAX_B64) {
    console.warn('PDF muy grande, se omite y se pide texto en servidor', payload.pdfBase64.length);
    payload.pdfBase64 = '';
    payload.generarPdfTexto = true;
    payload.pdfOrigen = 'omit_size';
  }
  if (payload.excelBase64 && payload.excelBase64.length > MAX_B64) {
    console.warn('Excel muy grande, se omite', payload.excelBase64.length);
    payload.excelBase64 = '';
  }
  if (payload.wordBase64 && payload.wordBase64.length > MAX_B64) {
    payload.wordBase64 = '';
  }

  let pdfInfo = null;
  let sheetInfo = null;
  let sheetError = '';
  let wordInfo = null;
  let excelInfo = null;

  try {
    // —— 1) Archivos en Drive (PDF, Excel, Word) en un POST; si falla, por partes ——
    setEstadoSheets('☁ Subiendo PDF y Excel a Google Drive…', false);
    const filesPayload = {
      action: 'save_files',
      pdfName: payload.pdfName,
      wordName: payload.wordName || nombreArchivoWordFicha(),
      excelName: payload.excelName || nombreArchivoFicha('xlsx'),
      meta: payload.meta,
      headers: payload.headers,
      values: payload.values,
      pdfBase64: payload.pdfBase64 || '',
      wordBase64: payload.wordBase64 || '',
      excelBase64: payload.excelBase64 || '',
      pdfOrigen: payload.pdfOrigen || '',
      skipPdf: !payload.pdfBase64
    };

    function applyFilesResult(r) {
      if (!r || !r.data || !r.data.ok) return false;
      if (r.data.pdfUrl) {
        pdfInfo = r.data;
        if (r.data.pdfOrigen === 'ficha_impresa' || payload.pdfBase64) {
          pdfInfo.esFichaImpresa = true;
        }
      }
      if (r.data.wordUrl) {
        wordInfo = {
          wordUrl: r.data.wordUrl,
          wordName: r.data.wordName,
          wordId: r.data.wordId
        };
      }
      if (r.data.excelUrl) {
        excelInfo = {
          excelUrl: r.data.excelUrl,
          excelName: r.data.excelName,
          excelId: r.data.excelId
        };
      }
      return !!(pdfInfo || wordInfo || excelInfo);
    }

    let rFiles;
    try {
      rFiles = await postJson(filesPayload, 180000);
    } catch (eFiles) {
      console.warn('save_files network', eFiles);
      rFiles = null;
    }

    if (rFiles && rFiles.authError) {
      setEstadoSheets(mensajeErrorAuthAppsScript(), true);
      hablar('Google pide permiso. Configure la aplicación web como pública.');
      if (btn) btn.disabled = false;
      return { ok: false, error: 'auth_401' };
    }

    if (!applyFilesResult(rFiles)) {
      console.warn('save_files falló, reintentando por partes', rFiles && (rFiles.text || '').slice(0, 200));
      // PDF solo
      if (payload.pdfBase64) {
        try {
          const rPdf = await postJson({
            action: 'save_pdf_only',
            pdfName: payload.pdfName,
            meta: payload.meta,
            pdfBase64: payload.pdfBase64,
            pdfOrigen: payload.pdfOrigen || ''
          }, 120000);
          applyFilesResult(rPdf);
        } catch (e1) { console.warn(e1); }
      }
      // Excel solo
      if (payload.excelBase64) {
        try {
          const rEx = await postJson({
            action: 'save_files',
            skipPdf: true,
            excelName: payload.excelName || nombreArchivoFicha('xlsx'),
            excelBase64: payload.excelBase64,
            meta: payload.meta
          }, 120000);
          applyFilesResult(rEx);
        } catch (e2) { console.warn(e2); }
      }
      // Si no hubo PDF base64, pedir PDF de texto en servidor
      if (!pdfInfo) {
        try {
          const rTxt = await postJson({
            action: 'save_pdf_only',
            pdfName: payload.pdfName,
            meta: payload.meta,
            headers: payload.headers,
            values: payload.values,
            skipPdf: false
          }, 60000);
          applyFilesResult(rTxt);
        } catch (e3) { console.warn(e3); }
      }
    }

    // —— 2) Fila en la hoja (sin archivos pesados) ——
    setEstadoSheets(
      (pdfInfo || wordInfo || excelInfo)
        ? '✓ Archivos en Drive. Guardando fila en la hoja…'
        : '☁ Guardando fila en la hoja…',
      false
    );
    const sheetPayload = {
      action: 'append_row_only',
      headers: payload.headers,
      values: payload.values,
      meta: payload.meta,
      pdfName: payload.pdfName,
      pdfUrlPrevio: pdfInfo ? (pdfInfo.pdfUrl || '') : '',
      pdfNamePrevio: pdfInfo ? (pdfInfo.pdfName || '') : '',
      wordUrlPrevio: wordInfo ? (wordInfo.wordUrl || '') : '',
      wordNamePrevio: wordInfo ? (wordInfo.wordName || '') : '',
      excelUrlPrevio: excelInfo ? (excelInfo.excelUrl || '') : '',
      excelNamePrevio: excelInfo ? (excelInfo.excelName || '') : ''
    };
    let rSheet;
    try {
      rSheet = await postJson(sheetPayload, 60000);
    } catch (eSheet) {
      rSheet = { data: null, text: String(eSheet && eSheet.message ? eSheet.message : eSheet), authError: false };
    }
    if (rSheet.authError) {
      setEstadoSheets(mensajeErrorAuthAppsScript(), true);
      if (btn) btn.disabled = false;
      return { ok: false, error: 'auth_401' };
    }
    if (rSheet.data && rSheet.data.ok) {
      sheetInfo = rSheet.data;
      if (rSheet.data.pdfUrl) pdfInfo = Object.assign({}, pdfInfo || {}, rSheet.data);
      if (rSheet.data.excelUrl) {
        excelInfo = Object.assign({}, excelInfo || {}, {
          excelUrl: rSheet.data.excelUrl,
          excelName: rSheet.data.excelName
        });
      }
    } else {
      sheetError = (rSheet.data && rSheet.data.error) || rSheet.text || 'Error al escribir la hoja';
      console.warn('append falló', sheetError);
      const ver = (rSheet.data && rSheet.data.version) || '';
      if (!ver || /v4/i.test(ver) || /columnas/i.test(sheetError)) {
        mostrarAvisoScriptViejo(true);
        sheetError =
          'Script de Google desactualizado (' + (ver || '?') + '). ' +
          'Pegue Codigo.gs, Implementar → Nueva versión. Detalle: ' +
          String(sheetError).slice(0, 100);
      }
    }

    const pdfOk = !!(pdfInfo && pdfInfo.pdfUrl);
    const wordOk = !!(wordInfo && wordInfo.wordUrl);
    const excelOk = !!(excelInfo && excelInfo.excelUrl);
    const sheetOk = !!(sheetInfo && sheetInfo.ok !== false && (sheetInfo.row || sheetInfo.message));
    const verSheet = (sheetInfo && sheetInfo.version) || (rSheet && rSheet.data && rSheet.data.version) || '';
    if (verSheet && /v(8|9|10)/i.test(verSheet)) mostrarAvisoScriptViejo(false);

    if (!pdfOk && !wordOk && !excelOk && !sheetOk) {
      setEstadoSheets('No se pudo guardar en Google: ' + (sheetError || 'sin respuesta'), true);
      hablar('No se pudo guardar en Google. Revise internet e intente de nuevo.');
      if (btn) btn.disabled = false;
      return { ok: false, error: sheetError || 'fallo_total' };
    }

    let estado = '';
    if (sheetOk) {
      estado += '✓ Ficha en la hoja' + (sheetInfo.row ? ' (fila ' + sheetInfo.row + ', pestaña ' + (sheetInfo.sheet || 'Fichas') + ')' : '') + '. ';
    } else {
      estado += '⚠ Fila en hoja NO guardada. ' + String(sheetError).slice(0, 120) + ' ';
    }
    if (pdfOk) estado += '✓ PDF: ' + (pdfInfo.pdfName || 'sí') + '. ';
    else estado += '⚠ PDF no creado. ';
    if (excelOk) estado += '✓ Excel: ' + (excelInfo.excelName || 'sí') + '. ';
    else estado += '⚠ Excel no creado. ';
    if (wordOk) estado += '✓ Word: ' + (wordInfo.wordName || 'sí') + '.';
    else estado += '⚠ Word no creado.';
    setEstadoSheets(estado, !(sheetOk && (pdfOk || excelOk || wordOk)));

    let voz = '';
    if (sheetOk && (pdfOk || wordOk || excelOk)) {
      voz = 'Listo. Datos en la hoja';
      if (pdfOk) voz += ', PDF en Drive';
      if (excelOk) voz += ', Excel en Drive';
      if (wordOk) voz += ' y Word en Drive';
      voz += '.';
    } else if (pdfOk || wordOk || excelOk) {
      voz = 'Archivos guardados en Drive';
      if (pdfOk) voz += ' (PDF)';
      if (excelOk) voz += ' (Excel)';
      if (wordOk) voz += ' (Word)';
      voz += ', pero la fila de la hoja falló. Actualice el script de Google a la versión 10.';
    } else if (sheetOk) {
      voz = 'La ficha se guardó en la hoja, pero no se crearon PDF ni Excel.';
    } else {
      voz = 'No se pudo guardar. Revise la configuración de Google.';
    }
    if (sheetInfo && sheetInfo.row) voz += ' Fila ' + sheetInfo.row + '.';
    hablar(voz);

    try {
      localStorage.setItem('fichaUNCP_lastSheetsOk', new Date().toISOString());
      localStorage.setItem('fichaUNCP_lastSheetsDni', String((respuestas && respuestas.dni) || ''));
      if (pdfInfo && pdfInfo.pdfUrl) localStorage.setItem('fichaUNCP_lastPdfUrl', pdfInfo.pdfUrl);
      if (wordInfo && wordInfo.wordUrl) localStorage.setItem('fichaUNCP_lastWordUrl', wordInfo.wordUrl);
      if (excelInfo && excelInfo.excelUrl) localStorage.setItem('fichaUNCP_lastExcelUrl', excelInfo.excelUrl);
    } catch (e3) { /* ignore */ }

    if (btn) btn.disabled = false;
    const msgEl = document.getElementById('mensaje-resumen');
    if (msgEl) {
      msgEl.textContent = (pdfOk || excelOk)
        ? (sheetOk ? '✅ Ficha + PDF/Excel guardados' : '✅ Archivos en Drive (hoja con aviso)')
        : (sheetOk ? '✅ Ficha en hoja (archivos pendientes)' : '⚠ Error al guardar');
    }
    const est = document.getElementById('estado-sheets');
    if (est) {
      let links = '';
      if (pdfOk) {
        links += ' <a href="' + esc(pdfInfo.pdfUrl) + '" target="_blank" rel="noopener" style="color:var(--info);">Abrir PDF</a>';
      }
      if (excelOk) {
        links += ' · <a href="' + esc(excelInfo.excelUrl) + '" target="_blank" rel="noopener" style="color:var(--info);">Abrir Excel</a>';
      }
      if (wordOk) {
        links += ' · <a href="' + esc(wordInfo.wordUrl) + '" target="_blank" rel="noopener" style="color:var(--info);">Abrir Word</a>';
      }
      links += ' · <a href="' + esc(SHEETS_DRIVE_FOLDER_URL) + '" target="_blank" rel="noopener" style="color:var(--info);">Carpeta</a>';
      est.innerHTML = esc(estado) + links;
    }
    if (sheetOk || pdfOk || wordOk || excelOk) {
      if (sheetOk) mostrarAvisoScriptViejo(false);
    }
    return { ok: pdfOk || wordOk || excelOk || sheetOk, pdfInfo, wordInfo, excelInfo, sheetInfo, sheetError };
  } catch (err) {
    console.error('Sheets/Drive', err);
    const m = String(err && err.message ? err.message : err);
    let msg = 'Error al guardar: ' + m;
    if (/Failed to fetch|NetworkError|Load failed|abort/i.test(m)) {
      msg =
        'No se pudo contactar a Google (Failed to fetch). ' +
        'Abra su URL /exec en el navegador: debe verse texto JSON con "ok":true. ' +
        'Si ve login de Google, en Apps Script ponga «Quién tiene acceso: Cualquier persona» ' +
        'y cree una Nueva versión de la implementación.';
    }
    setEstadoSheets(msg, true);
    hablar('No se pudo guardar en Google. Revise la implementación del script.');
    if (btn) btn.disabled = false;
    return { ok: false, error: m };
  }
}

// ═══════════════════════════════════════════
// UTILS
// ═══════════════════════════════════════════
function esc(s) {
  return String(s ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

document.addEventListener('keydown', (e) => {
  const enEntrevista = document.getElementById('pantalla-entrevista')?.classList.contains('activa');
  if (!enEntrevista) return;
  if (e.altKey && e.key === 'm') { e.preventDefault(); toggleMicrofono(); }
  if (e.altKey && e.key === 'Enter') { e.preventDefault(); confirmarYSiguiente(); }
  if (e.altKey && e.key === 'r') { e.preventDefault(); repetirPregunta(); }
});

function precargarVoces(cb) {
  if (!('speechSynthesis' in window)) {
    if (cb) cb();
    return;
  }
  const ready = () => {
    const v = obtenerVoces();
    if (v.length) vocesListas = true;
    if (cb) cb();
  };
  // Chrome carga voces de forma asíncrona
  const existentes = obtenerVoces();
  if (existentes.length) {
    ready();
  } else {
    speechSynthesis.onvoiceschanged = () => {
      ready();
    };
    // Forzar carga
    try { speechSynthesis.getVoices(); } catch (e) { /* ignore */ }
    setTimeout(ready, 600);
  }
}

window.addEventListener('load', () => {
  // Narrador OFF por defecto (nunca suena al abrir la app)
  usuarioEligioVoz = false;
  try { localStorage.setItem(LECTURA_VOZ_KEY, '0'); } catch (e) { /* ignore */ }
  setLecturaVoz(false, { persistir: false });
  detenerHabla();

  precargarVoces(() => actualizarAvisoVozUI());
  actualizarAvisoVozUI();
  actualizarEstadoSheetsUI();
  actualizarUIAdmin();

  const raw = localStorage.getItem('fichaUNCP_v3') || localStorage.getItem('fichaUNCP_v2');
  if (raw) {
    const btn = document.getElementById('btn-cargar-inicio');
    if (btn) btn.style.display = 'block';
  }

  // Atajo admin: Alt+Shift+A
  document.addEventListener('keydown', function (ev) {
    if (ev.altKey && ev.shiftKey && (ev.key === 'A' || ev.key === 'a')) {
      ev.preventDefault();
      if (esModoAdmin()) salirModoAdmin();
      else pedirClaveAdmin();
    }
  });

  // Sin audio al cargar: mensaje solo en pantalla
  mostrarMensaje(
    'Listo. Por defecto no hay narrador. Pulse «Comenzar sin voz» o, si lo desea, «Comenzar con voz».'
  );
});
