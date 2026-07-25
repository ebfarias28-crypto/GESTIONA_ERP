import React, { useState, useMemo, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import { supabase } from './supabaseClient';
import {
  LayoutDashboard, ClipboardList, Users, UserCog, Package, FileText, Wallet,
  Plus, Trash2, X, MapPin, Phone, AlertTriangle, Cable, BookOpen, FolderKanban,
  Banknote, BarChart3, Download, ArrowUpCircle, ArrowDownCircle,
  Building2, Truck, UserCheck, Calculator, ChevronDown, Pencil, Check, Fuel, ShieldAlert, Bell
} from 'lucide-react';

const C = {
  navy: '#0B1622',
  navyLight: '#16283B',
  bg: '#EEF1F3',
  card: '#FFFFFF',
  cyan: '#00C2B8',
  cyanDark: '#039087',
  amber: '#F5A623',
  coral: '#E5484D',
  slate: '#1F2933',
  muted: '#64748B',
  border: '#E2E5E9',
};

let MONEDA_SIMBOLO = '$'; // se actualiza desde las Preferencias del usuario en sesión
const fmt = (n) => MONEDA_SIMBOLO + Number(n || 0).toLocaleString('es-MX', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

const nextId = (prefix, list, pad = 3) => {
  const nums = list.map((x) => parseInt(String(x.id).replace(prefix, ''), 10) || 0);
  const next = (nums.length ? Math.max(...nums) : 0) + 1;
  return prefix + String(next).padStart(pad, '0');
};

const nextFolio = (list) => {
  const nums = list.map((o) => parseInt(o.id.split('-').pop(), 10) || 0);
  const next = (nums.length ? Math.max(...nums) : 0) + 1;
  return 'OT-2607-' + String(next).padStart(3, '0');
};

const INIT_CLIENTS = [
  { id: 'C001', nombre: 'Comercial Andina S.A.', telefono: '322-455-0110', direccion: 'Av. Reforma 245, Col. Centro' },
  { id: 'C002', nombre: 'Residencial Las Palmas', telefono: '322-455-0187', direccion: 'Calle Palmas 12' },
  { id: 'C003', nombre: 'Farmacias del Valle', telefono: '322-455-0223', direccion: 'Blvd. Independencia 890' },
  { id: 'C004', nombre: 'María Fernanda Ruiz', telefono: '322-455-0341', direccion: 'Priv. Jazmines 45' },
  { id: 'C005', nombre: 'Hotel Bahía Azul', telefono: '322-455-0456', direccion: 'Costera Sur km 4' },
];

const SEED_MATERIALS = [
  { id: 'M001', nombre: 'Roseta óptica', unidad: 'pza', costoUnitario: 45, stock: 120 },
  { id: 'M002', nombre: 'Conector SC/APC', unidad: 'pza', costoUnitario: 18, stock: 60 },
  { id: 'M003', nombre: 'Cable drop 3mm (rollo 100m)', unidad: 'rollo', costoUnitario: 520, stock: 15 },
  { id: 'M004', nombre: 'ONT / Router GPON', unidad: 'pza', costoUnitario: 780, stock: 30 },
  { id: 'M005', nombre: 'Splitter 1x8', unidad: 'pza', costoUnitario: 210, stock: 6 },
  { id: 'M006', nombre: 'Manga de empalme', unidad: 'pza', costoUnitario: 165, stock: 14 },
];

const INIT_PROJECTS = [
  { id: 'P001', nombre: 'Fraccionamiento Las Palmas — Fase 1', clienteId: 'C002', fechaInicio: '2026-06-15', presupuesto: 15000, estado: 'activo' },
  { id: 'P002', nombre: 'Corporativo Andina — Backbone interno', clienteId: 'C001', fechaInicio: '2026-07-01', presupuesto: 6000, estado: 'activo' },
];

const SEED_ORDERS = [
  { id: 'OT-2607-001', proyectoId: 'P002', centroCostoId: 'CC-IMP', clienteId: 'C001', tecnicoId: 'T002', tipoServicio: 'Instalación empresarial', direccion: 'Av. Reforma 245, Col. Centro', fecha: '2026-07-02', estado: 'completada', materialesUsados: [{ materialId: 'M001', cantidad: 2 }, { materialId: 'M002', cantidad: 4 }, { materialId: 'M004', cantidad: 1 }], costoManoObra: 650, precioCliente: 2400, estadoPago: 'pagada', estadoPagoTecnico: 'pagado' },
  { id: 'OT-2607-002', proyectoId: 'P001', centroCostoId: 'CC-EXC', clienteId: 'C002', tecnicoId: 'T001', tipoServicio: 'Instalación residencial', direccion: 'Calle Palmas 12', fecha: '2026-07-03', estado: 'completada', materialesUsados: [{ materialId: 'M001', cantidad: 1 }, { materialId: 'M004', cantidad: 1 }], costoManoObra: 450, precioCliente: 1200, estadoPago: 'pendiente', estadoPagoTecnico: 'pendiente' },
  { id: 'OT-2607-003', proyectoId: null, centroCostoId: 'CC-IMP', clienteId: 'C003', tecnicoId: 'T003', tipoServicio: 'Empalme de fibra', direccion: 'Blvd. Independencia 890', fecha: '2026-07-05', estado: 'completada', materialesUsados: [{ materialId: 'M006', cantidad: 2 }, { materialId: 'M002', cantidad: 6 }], costoManoObra: 550, precioCliente: 1600, estadoPago: 'pagada', estadoPagoTecnico: 'pendiente' },
  { id: 'OT-2607-004', proyectoId: null, centroCostoId: 'CC-EXC', clienteId: 'C004', tecnicoId: 'T001', tipoServicio: 'Instalación residencial', direccion: 'Priv. Jazmines 45', fecha: '2026-07-08', estado: 'en_proceso', materialesUsados: [{ materialId: 'M001', cantidad: 1 }], costoManoObra: 450, precioCliente: 1150, estadoPago: 'pendiente', estadoPagoTecnico: 'pendiente' },
  { id: 'OT-2607-005', proyectoId: null, centroCostoId: 'CC-CRE', clienteId: 'C005', tecnicoId: 'T002', tipoServicio: 'Instalación empresarial', direccion: 'Costera Sur km 4', fecha: '2026-07-09', estado: 'pendiente', materialesUsados: [], costoManoObra: 650, precioCliente: 3200, estadoPago: 'pendiente', estadoPagoTecnico: 'pendiente' },
  { id: 'OT-2607-006', proyectoId: 'P002', centroCostoId: 'CC-IMP', clienteId: 'C001', tecnicoId: 'T003', tipoServicio: 'Mantenimiento / reparación', direccion: 'Av. Reforma 245, Col. Centro', fecha: '2026-07-10', estado: 'pendiente', materialesUsados: [], costoManoObra: 400, precioCliente: 800, estadoPago: 'pendiente', estadoPagoTecnico: 'pendiente' },
  { id: 'OT-2607-007', proyectoId: null, centroCostoId: 'CC-EXC', clienteId: 'C002', tecnicoId: 'T001', tipoServicio: 'Instalación residencial', direccion: 'Calle Palmas 12', fecha: '2026-06-28', estado: 'cancelada', materialesUsados: [], costoManoObra: 0, precioCliente: 0, estadoPago: 'pendiente', estadoPagoTecnico: 'pendiente' },
  { id: 'OT-2607-008', proyectoId: null, centroCostoId: 'CC-CRE', clienteId: 'C003', tecnicoId: 'T002', tipoServicio: 'Instalación empresarial', direccion: 'Blvd. Independencia 890', fecha: '2026-07-06', estado: 'completada', materialesUsados: [{ materialId: 'M003', cantidad: 1 }, { materialId: 'M005', cantidad: 1 }, { materialId: 'M004', cantidad: 2 }], costoManoObra: 650, precioCliente: 4100, estadoPago: 'pagada', estadoPagoTecnico: 'pagado' },
];

// ---------- Cuentas por pagar (proveedores) ----------
const SEED_PROVEEDORES = [
  { id: 'PR001', nombre: 'FiberTec Distribuciones', contacto: 'Jorge Salas', telefono: '322-200-3301' },
  { id: 'PR002', nombre: 'Ferretería Industrial del Pacífico', contacto: 'Ana Ibarra', telefono: '322-200-3302' },
  { id: 'PR003', nombre: 'Combustibles del Valle', contacto: 'Mostrador', telefono: '322-200-3303' },
];
const SEED_CXP = [
  { id: 'CP001', proveedorId: 'PR001', concepto: 'Compra de ONTs y conectores', fecha: '2026-07-01', vencimiento: '2026-07-31', monto: 8200, estado: 'pendiente' },
  { id: 'CP002', proveedorId: 'PR002', concepto: 'Herramienta menor', fecha: '2026-06-20', vencimiento: '2026-07-05', monto: 1350, estado: 'pagada' },
  { id: 'CP003', proveedorId: 'PR003', concepto: 'Vale de combustible mensual', fecha: '2026-07-05', vencimiento: '2026-07-20', monto: 2200, estado: 'pendiente' },
];

// ---------- Activos fijos ----------
const SEED_ACTIVOS = [
  { id: 'AF001', nombre: 'Camioneta Nissan NP300', categoria: 'Vehículo', fechaCompra: '2024-03-10', costo: 380000, valorResidual: 60000, vidaUtilAnios: 5, responsable: 'T001', centroCostoId: 'CC-IMP' },
  { id: 'AF002', nombre: 'Fusionadora de fibra Fujikura', categoria: 'Equipo', fechaCompra: '2025-01-15', costo: 95000, valorResidual: 5000, vidaUtilAnios: 4, responsable: 'T003', centroCostoId: 'CC-IMP' },
  { id: 'AF003', nombre: 'Camioneta Toyota Hilux', categoria: 'Vehículo', fechaCompra: '2023-08-01', costo: 420000, valorResidual: 70000, vidaUtilAnios: 5, responsable: 'T002', centroCostoId: 'CC-EXC' },
  { id: 'AF004', nombre: 'Kit de herramientas de instalación', categoria: 'Herramienta', fechaCompra: '2025-11-01', costo: 18000, valorResidual: 1000, vidaUtilAnios: 3, responsable: 'T001', centroCostoId: 'CC-CRE' },
];

// ---------- Recursos humanos ----------
const SEED_EMPLEADOS = [
  { id: 'E001', nombre: 'Patricia Ley', rol: 'Administrativo', puesto: 'Administración y ventas', telefono: '322-300-1101', fechaIngreso: '2023-02-01', salarioMensual: 9000, reportaA: null },
  { id: 'E002', nombre: 'Roberto Aguilar', rol: 'Administrativo', puesto: 'Supervisor de operaciones', telefono: '322-300-1102', fechaIngreso: '2022-09-15', salarioMensual: 12500, reportaA: 'E001' },
  { id: 'T001', nombre: 'Carlos Medina', rol: 'Técnico', puesto: 'Instalación residencial', telefono: '322-100-2201', fechaIngreso: '2023-05-10', salarioMensual: 9500, reportaA: 'E002' },
  { id: 'T002', nombre: 'Luis Ángel Torres', rol: 'Técnico', puesto: 'Instalación empresarial', telefono: '322-100-2202', fechaIngreso: '2022-11-20', salarioMensual: 10500, reportaA: 'E002' },
  { id: 'T003', nombre: 'Diana Gómez', rol: 'Técnico', puesto: 'Empalmes y fusión', telefono: '322-100-2203', fechaIngreso: '2024-01-08', salarioMensual: 9800, reportaA: 'E002' },
];
const SEED_NOMINA = [
  { id: 'NM001', empleadoId: 'E001', periodo: 'Junio 2026', fechaPago: '2026-06-30', monto: 9000, estado: 'pagado' },
  { id: 'NM002', empleadoId: 'E002', periodo: 'Junio 2026', fechaPago: '2026-06-30', monto: 12500, estado: 'pagado' },
  { id: 'NM003', empleadoId: 'E001', periodo: 'Julio 2026', fechaPago: '2026-07-31', monto: 9000, estado: 'pendiente' },
  { id: 'NM004', empleadoId: 'E002', periodo: 'Julio 2026', fechaPago: '2026-07-31', monto: 12500, estado: 'pendiente' },
];

// ---------- Contabilidad general ----------
const ER_GRUPOS = [
  { key: 'ingresos_operativos', label: 'Ingresos por instalaciones' },
  { key: 'costo_materiales', label: 'Costo de materiales' },
  { key: 'costo_mano_obra', label: 'Costo de mano de obra' },
  { key: 'costo_depreciacion', label: 'Costo de depreciación' },
  { key: 'otros_ingresos', label: 'Otros ingresos' },
  { key: 'gastos_generales', label: 'Gastos generales' },
];

const BG_SUBTIPOS = [
  { key: 'activo_corriente', label: 'Activo corriente', tipo: 'Activo' },
  { key: 'activo_no_corriente', label: 'Activo no corriente', tipo: 'Activo' },
  { key: 'pasivo_corriente', label: 'Pasivo corriente', tipo: 'Pasivo' },
  { key: 'pasivo_no_corriente', label: 'Pasivo no corriente', tipo: 'Pasivo' },
];

const SEED_PLAN_CUENTAS = [
  { codigo: '1105', nombre: 'Caja', tipo: 'Activo', subtipo: 'activo_corriente', grupo: null },
  { codigo: '1110', nombre: 'Bancos', tipo: 'Activo', subtipo: 'activo_corriente', grupo: null },
  { codigo: '1305', nombre: 'Cuentas por cobrar clientes', tipo: 'Activo', subtipo: 'activo_corriente', grupo: null },
  { codigo: '1310', nombre: 'Anticipos a empleados', tipo: 'Activo', subtipo: 'activo_corriente', grupo: null },
  { codigo: '1315', nombre: 'Préstamos a empleados', tipo: 'Activo', subtipo: 'activo_corriente', grupo: null },
  { codigo: '1435', nombre: 'Inventario de materiales', tipo: 'Activo', subtipo: 'activo_corriente', grupo: null },
  { codigo: '1540', nombre: 'Activos fijos', tipo: 'Activo', subtipo: 'activo_no_corriente', grupo: null },
  { codigo: '1541', nombre: 'Depreciación acumulada — activos fijos', tipo: 'Activo', subtipo: 'activo_no_corriente', grupo: null },
  { codigo: '2205', nombre: 'Cuentas por pagar proveedores', tipo: 'Pasivo', subtipo: 'pasivo_corriente', grupo: null },
  { codigo: '2510', nombre: 'Nómina y mano de obra por pagar', tipo: 'Pasivo', subtipo: 'pasivo_corriente', grupo: null },
  { codigo: '3105', nombre: 'Capital social', tipo: 'Capital', subtipo: null, grupo: null },
  { codigo: '4135', nombre: 'Ingresos por instalaciones', tipo: 'Ingreso', subtipo: null, grupo: 'ingresos_operativos' },
  { codigo: '4200', nombre: 'Otros ingresos', tipo: 'Ingreso', subtipo: null, grupo: 'otros_ingresos' },
  { codigo: '6135', nombre: 'Costo de materiales', tipo: 'Costo', subtipo: null, grupo: 'costo_materiales' },
  { codigo: '6140', nombre: 'Costo de mano de obra', tipo: 'Costo', subtipo: null, grupo: 'costo_mano_obra' },
  { codigo: '6310', nombre: 'Costo de depreciación', tipo: 'Costo', subtipo: null, grupo: 'costo_depreciacion' },
  { codigo: '5310', nombre: 'Gasto de depreciación (administrativo)', tipo: 'Gasto', subtipo: null, grupo: 'gastos_generales' },
  { codigo: '5195', nombre: 'Gastos generales', tipo: 'Gasto', subtipo: null, grupo: 'gastos_generales' },
  { codigo: '5105', nombre: 'Gastos de nómina administrativa', tipo: 'Gasto', subtipo: null, grupo: 'gastos_generales' },
];
const SEED_GASTO_CATEGORIAS = [
  { nombre: 'Renta de local/bodega', cuenta: '5195' },
  { nombre: 'Combustible y transporte', cuenta: '5195' },
  { nombre: 'Bonos y compensaciones extra', cuenta: '5195' },
  { nombre: 'Servicios (luz/internet/agua)', cuenta: '5195' },
  { nombre: 'Herramientas y equipo', cuenta: '5195' },
  { nombre: 'Seguros', cuenta: '5195' },
  { nombre: 'Impuestos', cuenta: '5195' },
  { nombre: 'Otro gasto', cuenta: '5195' },
];
const SEED_INGRESO_CATEGORIAS = [
  { nombre: 'Venta de equipo/material sobrante', cuenta: '4200' },
  { nombre: 'Renta de herramienta', cuenta: '4200' },
  { nombre: 'Rendimientos financieros', cuenta: '4200' },
  { nombre: 'Otro ingreso', cuenta: '4200' },
];

const SEED_MOVIMIENTOS = [
  { id: 'MV001', tipo: 'gasto', categoria: 'Renta de local/bodega', descripcion: 'Renta de bodega julio', fecha: '2026-07-01', monto: 4500 },
  { id: 'MV002', tipo: 'gasto', categoria: 'Combustible y transporte', descripcion: 'Gasolina camionetas de instalación', fecha: '2026-07-03', monto: 1850 },
  { id: 'MV003', tipo: 'gasto', categoria: 'Bonos y compensaciones extra', descripcion: 'Bono de productividad — equipo administrativo', fecha: '2026-07-05', monto: 1500 },
  { id: 'MV004', tipo: 'gasto', categoria: 'Servicios (luz/internet/agua)', descripcion: 'Luz e internet de oficina', fecha: '2026-07-06', monto: 980 },
  { id: 'MV005', tipo: 'ingreso', categoria: 'Venta de equipo/material sobrante', descripcion: 'Venta de router dado de baja', fecha: '2026-07-07', monto: 650 },
  { id: 'MV006', tipo: 'gasto', categoria: 'Herramientas y equipo', descripcion: 'Fusionadora de fibra — mantenimiento', fecha: '2026-07-09', monto: 1200 },
];

// Construye el kardex inicial (costeo promedio ponderado) a partir del stock semilla
// y del consumo histórico de las órdenes semilla, dejando materiales y kardex consistentes.
// También devuelve el costo de materiales que consumió cada orden, para contabilizarlo igual en el libro diario.
function buildInitialKardex(materials, orders) {
  const state = {};
  materials.forEach((m) => { state[m.id] = { stock: m.stock, costo: m.costoUnitario }; });
  const log = [];
  const orderMaterialCost = {};
  let seq = 1;
  materials.forEach((m) => {
    log.push({
      id: 'K' + String(seq++).padStart(4, '0'), materialId: m.id, fecha: '2026-06-01', tipo: 'entrada',
      cantidad: m.stock, costoUnitario: m.costoUnitario, referencia: 'Stock inicial',
      saldoCantidad: m.stock, saldoValor: m.stock * m.costoUnitario,
    });
  });
  const sorted = [...orders].filter((o) => o.estado !== 'cancelada').sort((a, b) => a.fecha.localeCompare(b.fecha));
  sorted.forEach((o) => {
    o.materialesUsados.forEach((mu) => {
      const s = state[mu.materialId];
      if (!s) return;
      s.stock -= mu.cantidad;
      orderMaterialCost[o.id] = (orderMaterialCost[o.id] || 0) + mu.cantidad * s.costo;
      log.push({
        id: 'K' + String(seq++).padStart(4, '0'), materialId: mu.materialId, fecha: o.fecha, tipo: 'salida',
        cantidad: mu.cantidad, costoUnitario: s.costo, referencia: `Orden ${o.id}`,
        saldoCantidad: s.stock, saldoValor: s.stock * s.costo,
      });
    });
  });
  const finalMaterials = materials.map((m) => ({ ...m, stock: state[m.id].stock, costoUnitario: state[m.id].costo }));
  return { log, finalMaterials, orderMaterialCost };
}

const { log: SEED_KARDEX, finalMaterials: INIT_MATERIALS, orderMaterialCost: SEED_ORDER_MATERIAL_COST } = buildInitialKardex(SEED_MATERIALS, SEED_ORDERS);

// Genera TODO el libro diario inicial a partir de los mismos datos semilla que usan los demás módulos,
// para que Contabilidad General cuadre desde el primer momento con Órdenes, Inventario, RH, CxP, Activos e Ingresos/Gastos.
function buildInitialAsientos({ orders, orderMaterialCost, materialsInitialValue, cxp, nomina, movimientos, activos }) {
  const log = [];
  let seq = 1;
  const post = (fecha, glosa, cuentaDebe, monto, cuentaHaber, centroCostoId = null) => {
    if (!monto) return;
    log.push({ id: 'AS' + String(seq++).padStart(3, '0'), fecha, glosa, cuentaDebe, montoDebe: monto, cuentaHaber, montoHaber: monto, centroCostoId });
  };

  post('2026-06-01', 'Apertura de inventario inicial', '1435', materialsInitialValue, '3105');
  activos.forEach((a) => post(a.fechaCompra, `Alta inicial de activo — ${a.nombre}`, '1540', a.costo, '3105', a.centroCostoId));

  [...orders].filter((o) => o.estado !== 'cancelada').sort((a, b) => a.fecha.localeCompare(b.fecha)).forEach((o) => {
    const costoMat = orderMaterialCost[o.id] || 0;
    post(o.fecha, `Consumo de materiales — orden ${o.id}`, '6135', costoMat, '1435', o.centroCostoId);
    post(o.fecha, `Mano de obra por pagar — orden ${o.id}`, '6140', o.costoManoObra, '2510', o.centroCostoId);
    if (o.estado === 'completada') {
      post(o.fecha, `Facturación — orden ${o.id}`, '1305', o.precioCliente, '4135', o.centroCostoId);
      if (o.estadoPago === 'pagada') post(o.fecha, `Cobro de cliente — orden ${o.id}`, '1110', o.precioCliente, '1305', o.centroCostoId);
    }
    if (o.estadoPagoTecnico === 'pagado') post(o.fecha, `Pago de mano de obra — orden ${o.id}`, '2510', o.costoManoObra, '1110', o.centroCostoId);
  });

  cxp.forEach((c) => {
    post(c.fecha, `Compra a crédito — ${c.concepto}`, '5195', c.monto, '2205', 'CC-ADM');
    if (c.estado === 'pagada') post(c.vencimiento, `Pago a proveedor — ${c.concepto}`, '2205', c.monto, '1110', 'CC-ADM');
  });

  nomina.forEach((n) => {
    post(n.fechaPago, `Nómina ${n.periodo}`, '5105', n.monto, '2510', 'CC-ADM');
    if (n.estado === 'pagado') post(n.fechaPago, `Pago de nómina — ${n.periodo}`, '2510', n.monto, '1110', 'CC-ADM');
  });

  movimientos.forEach((m) => {
    if (m.tipo === 'gasto') post(m.fecha, m.descripcion, '5195', m.monto, '1110', 'CC-ADM');
    else post(m.fecha, m.descripcion, '1110', m.monto, '4200', 'CC-ADM');
  });

  return log;
}

const SEED_ASIENTOS = buildInitialAsientos({
  orders: SEED_ORDERS,
  orderMaterialCost: SEED_ORDER_MATERIAL_COST,
  materialsInitialValue: SEED_MATERIALS.reduce((s, m) => s + m.stock * m.costoUnitario, 0),
  cxp: SEED_CXP,
  nomina: SEED_NOMINA,
  movimientos: SEED_MOVIMIENTOS,
  activos: SEED_ACTIVOS,
});

const SEED_CENTROS_COSTO = [
  { id: 'CC-CRE', nombre: 'Creación' },
  { id: 'CC-EXC', nombre: 'Excavación' },
  { id: 'CC-IMP', nombre: 'Implementación' },
  { id: 'CC-ADM', nombre: 'Administrativo' },
];

const CONTAB_VISTAS = [
  { key: 'balance_general', label: 'Balance general' },
  { key: 'estado_resultados', label: 'Estado de resultados' },
  { key: 'flujo_efectivo', label: 'Flujo de efectivo' },
  { key: 'balance_comprobacion', label: 'Balance de comprobación' },
  { key: 'libro_mayor', label: 'Libro mayor' },
];

// ---------- Multiempresa ----------
const GIROS_NEGOCIO = ['Servicios de fibra óptica', 'Construcción', 'Comercio', 'Transporte y logística', 'Manufactura', 'Servicios profesionales', 'Restaurantes y alimentos', 'Tecnología', 'Otro'];
const SEED_EMPRESAS = [
  { id: 'EMP001', nombre: 'BUILDERTECH SRL', giro: 'Servicios de fibra óptica', eslogan: 'Noi lo facciamo meglio', activa: true },
];

const CONFIG_SECCIONES = [
  { key: 'centros', label: 'Centros de costo' },
  { key: 'cat_gasto', label: 'Categorías de gasto' },
  { key: 'cat_ingreso', label: 'Categorías de ingreso' },
  { key: 'partidas', label: 'Partidas de trabajo / servicios' },
  { key: 'cierre', label: 'Cierre contable de periodo' },
];

// ---------- Roles y permisos configurables ----------
// Catálogo de módulos del sistema para la matriz de permisos
const MODULOS_PERMISOS = [
  { grupo: 'Indicadores KPI', items: [['dashboard', 'Panel general'], ['kpi_financieros', 'Índices financieros'], ['kpi_ordenes', 'Gestión de órdenes'], ['kpi_tecnicos', 'Rendimiento de técnicos'], ['kpi_inventario', 'Inventario y rotación'], ['kpi_cartera', 'Cartera y pagos']] },
  { grupo: 'Contabilidad General', items: [['contab_plan', 'Plan de Cuentas'], ['contab_comprobantes', 'Comprobantes'], ['contab_conciliacion', 'Conciliación bancaria'], ['contab_informes', 'Informes'], ['ingresos_gastos', 'Ingresos y gastos'], ['contab_config', 'Configuración']] },
  { grupo: 'Órdenes de trabajo', items: [['ordenes', 'Órdenes'], ['agenda', 'Agenda de técnicos'], ['proyectos', 'Proyectos'], ['ventas', 'Ventas (giro Comercio)']] },
  { grupo: 'Otros módulos', items: [['activos', 'Activos fijos'], ['depreciacion', 'Control de depreciación'], ['flota', 'Bitácora de flota'], ['facturacion', 'Facturación'], ['clientes', 'Clientes'], ['cxp', 'Cuentas por pagar'], ['proveedores', 'Proveedores'], ['inventario', 'Materiales'], ['kardex', 'Kardex'], ['rh', 'Empleados'], ['tecnicos', 'Técnicos'], ['bonos', 'Bonos por productividad'], ['reportes', 'Reportes']] },
];
const TODAS_LAS_CLAVES = MODULOS_PERMISOS.flatMap((g) => g.items.map(([k]) => k));
// Configuración inicial de cada rol (el Administrador siempre tiene acceso total y no es editable)
const DEFAULT_ROLES_CONFIG = {
  Administrador: null, // null = acceso total, no editable
  Contabilidad: ['dashboard', 'kpi_financieros', 'kpi_ordenes', 'kpi_cartera', 'contab_plan', 'contab_comprobantes', 'contab_conciliacion', 'contab_informes', 'ingresos_gastos', 'contab_config', 'facturacion', 'clientes', 'cxp', 'proveedores', 'reportes', 'activos', 'depreciacion', 'flota'],
  Operaciones: ['dashboard', 'kpi_ordenes', 'kpi_tecnicos', 'kpi_inventario', 'ordenes', 'agenda', 'proyectos', 'ventas', 'inventario', 'kardex', 'rh', 'tecnicos', 'bonos'],
  Tecnico: ['ordenes', 'agenda'],
};

const REPORT_TYPES = [
  { key: 'estado_resultados', label: 'Estado de resultados' },
  { key: 'ventas', label: 'Ventas (facturación)' },
  { key: 'costos', label: 'Costos por orden' },
  { key: 'gastos_categoria', label: 'Gastos por categoría' },
  { key: 'cxc', label: 'Cuentas por cobrar' },
  { key: 'cxp', label: 'Cuentas por pagar' },
  { key: 'inventario', label: 'Inventario valorizado' },
  { key: 'nomina', label: 'Nómina' },
  { key: 'activos', label: 'Activos fijos' },
  { key: 'proyectos', label: 'Utilidad por proyecto' },
  { key: 'tecnicos', label: 'Rentabilidad por técnico' },
  { key: 'centros_costo', label: 'Centros de costo (global y por centro)' },
];

const ESTADOS = {
  por_validar: { label: 'Por validar', bars: 0, color: '#8E6FCB' },
  pendiente: { label: 'Pendiente', bars: 1, color: C.amber },
  en_proceso: { label: 'En proceso', bars: 2, color: C.cyan },
  completada: { label: 'Completada', bars: 3, color: C.cyanDark },
  cancelada: { label: 'Cancelada', bars: 0, color: C.coral },
};

// Catálogo de partidas de trabajo (voci di lavorazione) — cada orden se desglosa en estas.
// Es un catálogo POR EMPRESA (editable en Configuración): esta es solo la semilla inicial de BuilderTech (fibra óptica).
const SEED_PARTIDAS_FTTH = [
  { id: 'PT01', nombre: 'Excavación / zanja', precioRef: 180 },
  { id: 'PT02', nombre: 'Tendido de fibra', precioRef: 220 },
  { id: 'PT03', nombre: 'Fusión / empalme', precioRef: 90 },
  { id: 'PT04', nombre: 'Instalación de ONT', precioRef: 120 },
  { id: 'PT05', nombre: 'Instalación de CPE / módem', precioRef: 80 },
  { id: 'PT06', nombre: 'Certificación / pruebas de señal', precioRef: 60 },
  { id: 'PT07', nombre: 'Reparación de avería', precioRef: 150 },
  { id: 'PT08', nombre: 'Traslado de punto', precioRef: 100 },
];
// Semilla genérica para una empresa nueva de cualquier giro — ella la adapta a su propio catálogo de partidas/servicios.
const SEED_PARTIDAS_GENERICO = [
  { id: 'PT01', nombre: 'Mano de obra', precioRef: 0 },
  { id: 'PT02', nombre: 'Materiales / insumos', precioRef: 0 },
  { id: 'PT03', nombre: 'Transporte / traslado', precioRef: 0 },
];

// Plantillas de correo — hoy solo se editan y previsualizan aquí; el envío real requiere backend (Supabase).
const EMAIL_VARIABLES = {
  agenda_squadra: ['{squadra}', '{fecha}', '{total_ordenes}', '{lista_ordenes}', '{empresa}'],
  cierre_cliente: ['{cliente}', '{folio}', '{fecha}', '{tecnico}', '{tipo_servicio}', '{empresa}'],
  vencimiento_flota: ['{vehiculo}', '{placa}', '{tipo_vencimiento}', '{fecha_vencimiento}', '{dias_restantes}'],
};
const SEED_EMAIL_TEMPLATES = {
  agenda_squadra: {
    nombre: 'Agenda diaria a la squadra',
    asunto: 'Agenda del {fecha} — Squadra {squadra}',
    cuerpo: 'Hola equipo {squadra},\n\nEsta es la agenda de hoy {fecha}: tienen {total_ordenes} orden(es) asignadas.\n\n{lista_ordenes}\n\nSaludos,\n{empresa}',
  },
  cierre_cliente: {
    nombre: 'Cierre de actividad al cliente',
    asunto: 'Trabajo completado — {empresa}',
    cuerpo: 'Estimado/a {cliente},\n\nLe confirmamos que el trabajo de {tipo_servicio} (folio {folio}) fue completado el {fecha} por nuestro técnico {tecnico}.\n\nGracias por su confianza.\n{empresa}',
  },
  vencimiento_flota: {
    nombre: 'Aviso de vencimiento de flota',
    asunto: 'Aviso: {tipo_vencimiento} de {vehiculo} vence pronto',
    cuerpo: 'El {tipo_vencimiento} del vehículo {vehiculo} (placa {placa}) vence el {fecha_vencimiento} (en {dias_restantes} días). Favor de gestionar la renovación.',
  },
};

// ---------- Idioma (Fase 1): navegación, títulos de módulo y textos de mayor tráfico ----------
// Patrón: la clave ES el texto en español; si el usuario elige italiano, se busca aquí. Si no está, se muestra el español tal cual.
const I18N_IT = {
  // Navegación
  'Indicadores KPI': 'Indicatori KPI', 'Panel general': 'Pannello generale', 'Índices financieros': 'Indici finanziari',
  'Gestión de órdenes': 'Gestione ordini', 'Rendimiento de técnicos': 'Rendimento dei tecnici', 'Inventario y rotación': 'Inventario e rotazione',
  'Cartera y pagos': 'Portafoglio e pagamenti', 'Contabilidad General': 'Contabilità Generale', 'Plan de Cuentas': 'Piano dei Conti',
  'Comprobantes': 'Comprovanti', 'Conciliación bancaria': 'Riconciliazione bancaria', 'Informes': 'Rapporti',
  'Ingresos y gastos': 'Entrate e uscite', 'Configuración': 'Configurazione', 'Órdenes de trabajo': 'Ordini di lavoro',
  'Órdenes': 'Ordini', 'Agenda de técnicos': 'Agenda tecnici', 'Proyectos': 'Progetti', 'Activos fijos': 'Beni strumentali',
  'Cuentas por cobrar': 'Crediti verso clienti', 'Facturas': 'Fatture', 'Clientes': 'Clienti', 'Cuentas por pagar': 'Debiti verso fornitori',
  'Inventarios': 'Magazzino', 'Materiales': 'Materiali', 'Kardex': 'Scheda di magazzino', 'Recursos humanos': 'Risorse umane',
  'Empleados': 'Dipendenti', 'Técnicos': 'Tecnici', 'Bonos por productividad': 'Bonus di produttività', 'Reportes': 'Rapporti',
  'Ventas': 'Vendite', 'Empresas': 'Aziende', 'Usuarios y roles': 'Utenti e ruoli',
  // Menú de usuario / login
  'Cambiar empresa': 'Cambia azienda', 'Cuenta': 'Account', 'Preferencias': 'Preferenze', 'Formato de correos': 'Modelli email',
  'Cerrar sesión': 'Esci', 'Iniciar sesión': 'Accedi', 'Usuario': 'Utente', 'Clave': 'Password', 'Entrar': 'Entra',
  // Títulos de módulo (h1)
  'Órdenes de trabajo · listado': 'Ordini di lavoro · elenco', 'Agenda semanal': 'Agenda settimanale',
  'Facturación a clientes': 'Fatturazione clienti',
  'Materiales y equipos': 'Materiali e attrezzature', 'Activos fijos y flota': 'Beni strumentali e flotta',
  'Mi cuenta': 'Il mio account',
  // Acciones comunes
  'Guardar': 'Salva', 'Guardar cambios': 'Salva modifiche', 'Cancelar': 'Annulla', 'Editar': 'Modifica', 'Eliminar': 'Elimina',
  'Nuevo': 'Nuovo', 'Nueva': 'Nuova', 'Cerrar': 'Chiudi', 'Buscar': 'Cerca', 'Exportar CSV': 'Esporta CSV',
  'Total': 'Totale', 'Estado': 'Stato', 'Fecha': 'Data', 'Cliente': 'Cliente', 'Técnico': 'Tecnico', 'Monto': 'Importo',
  // Títulos h1 adicionales
  'Inventario de materiales': 'Inventario materiali', 'Kardex de inventario': 'Scheda di magazzino',
  'Plan de cuentas': 'Piano dei conti', 'Ingresos y gastos generales': 'Entrate e uscite generali',
  // Fase 2: Agenda y Órdenes (contenido de formularios y tablas)
  'Semana del': 'Settimana dal', 'al': 'al', 'Semana anterior': 'Settimana precedente', 'Hoy': 'Oggi', 'Semana siguiente': 'Settimana successiva',
  'Todas': 'Tutte', 'Pendientes': 'In attesa', 'En proceso': 'In corso', 'Completadas': 'Completate', 'Canceladas': 'Annullate',
  'Tipo de servicio': 'Tipo di servizio', 'Dirección': 'Indirizzo', 'Fecha y hora': 'Data e ora', 'Centro de costo': 'Centro di costo',
  'Precio al cliente': 'Prezzo al cliente', 'Mano de obra': 'Manodopera', 'cobrado': 'incassato', 'por cobrar': 'da incassare',
  'pagada': 'pagata', 'por pagar': 'da pagare', 'Notas': 'Note', 'Historial de reprogramaciones': 'Storico riprogrammazioni',
  'Los montos y el estado se gestionan desde la pestaña Órdenes.': 'Gli importi e lo stato si gestiscono dalla scheda Ordini.',
  'Importar Excel': 'Importa Excel', 'Nueva orden': 'Nuovo ordine', 'Folio': 'Numero ordine', 'Proyecto': 'Progetto',
};

const ESTADOS_PROYECTO = {
  activo: { label: 'Activo', color: C.cyan },
  pausado: { label: 'Pausado', color: C.amber },
  completado: { label: 'Completado', color: C.cyanDark },
  cancelado: { label: 'Cancelado', color: C.coral },
};

// Gráfico de barras horizontal simple, sin dependencias externas — usado en los módulos de KPI
function BarChart({ title, data, colorFn, valueFmt }) {
  const max = Math.max(1, ...data.map((d) => Math.abs(d.value)));
  return (
    <div className="rounded-xl p-5" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
      {title && <div style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 600, fontSize: 14, color: C.slate, marginBottom: 12 }}>{title}</div>}
      {data.length === 0 && <div style={{ color: C.muted, fontSize: 12.5 }}>Sin datos en el periodo.</div>}
      {data.map((d, i) => (
        <div key={i} className="flex items-center gap-3 py-1.5">
          <span style={{ fontSize: 12, color: C.slate, width: 110, flexShrink: 0, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{d.label}</span>
          <div style={{ flex: 1, height: 12, backgroundColor: '#EEF1F4', borderRadius: 6, overflow: 'hidden' }}>
            <div style={{ width: `${(Math.abs(d.value) / max) * 100}%`, height: '100%', backgroundColor: colorFn ? colorFn(d) : C.cyanDark, borderRadius: 6, transition: 'width .3s' }} />
          </div>
          <span style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 12, width: 90, textAlign: 'right', flexShrink: 0 }}>{valueFmt ? valueFmt(d.value) : d.value}</span>
        </div>
      ))}
    </div>
  );
}

// Dona SVG de proporciones — usada para mostrar composición (ej. estados de órdenes, activo corriente vs no corriente)
function DonutChart({ title, data }) {
  const total = data.reduce((s, d) => s + d.value, 0);
  let acc = 0;
  const R = 42, C_ = 2 * Math.PI * R;
  return (
    <div className="rounded-xl p-5" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
      {title && <div style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 600, fontSize: 14, color: C.slate, marginBottom: 12 }}>{title}</div>}
      {total === 0 ? <div style={{ color: C.muted, fontSize: 12.5 }}>Sin datos en el periodo.</div> : (
        <div className="flex items-center gap-5">
          <svg width="110" height="110" viewBox="0 0 110 110">
            <g transform="translate(55,55) rotate(-90)">
              <circle r={R} fill="none" stroke="#EEF1F4" strokeWidth="16" />
              {data.map((d, i) => {
                const frac = d.value / total;
                const dash = frac * C_;
                const el = <circle key={i} r={R} fill="none" stroke={d.color} strokeWidth="16" strokeDasharray={`${dash} ${C_ - dash}`} strokeDashoffset={-acc} />;
                acc += dash;
                return el;
              })}
            </g>
          </svg>
          <div className="flex flex-col gap-1.5">
            {data.map((d, i) => (
              <div key={i} className="flex items-center gap-2" style={{ fontSize: 12 }}>
                <span style={{ width: 9, height: 9, borderRadius: 2, backgroundColor: d.color, display: 'inline-block' }} />
                <span style={{ color: C.slate }}>{d.label}</span>
                <span style={{ color: C.muted, fontFamily: 'JetBrains Mono, monospace' }}>{d.value} ({total > 0 ? Math.round((d.value / total) * 100) : 0}%)</span>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

// Paginación simple para tablas largas — evita renderizar cientos de filas de una vez
function usePaginado(lista, pageSize = 15) {
  const [pagina, setPagina] = useState(1);
  const totalPaginas = Math.max(1, Math.ceil(lista.length / pageSize));
  const paginaSegura = Math.min(pagina, totalPaginas);
  const inicio = (paginaSegura - 1) * pageSize;
  const paginaActual = lista.slice(inicio, inicio + pageSize);
  return { paginaActual, pagina: paginaSegura, totalPaginas, setPagina, total: lista.length, pageSize, inicio };
}
function Paginador({ p }) {
  if (p.total <= p.pageSize) return null;
  return (
    <div className="flex items-center justify-between px-4 py-3" style={{ borderTop: `1px solid ${C.border}` }}>
      <span style={{ fontSize: 11.5, color: C.muted }}>Mostrando {p.inicio + 1}–{Math.min(p.inicio + p.pageSize, p.total)} de {p.total}</span>
      <div className="flex items-center gap-2">
        <button onClick={() => p.setPagina(Math.max(1, p.pagina - 1))} disabled={p.pagina === 1} className="px-2.5 py-1 rounded-lg text-xs font-semibold" style={{ backgroundColor: '#EEF1F4', color: p.pagina === 1 ? C.muted : C.slate, opacity: p.pagina === 1 ? 0.5 : 1 }}>← Anterior</button>
        <span style={{ fontSize: 11.5, color: C.slate }}>Página {p.pagina} de {p.totalPaginas}</span>
        <button onClick={() => p.setPagina(Math.min(p.totalPaginas, p.pagina + 1))} disabled={p.pagina === p.totalPaginas} className="px-2.5 py-1 rounded-lg text-xs font-semibold" style={{ backgroundColor: '#EEF1F4', color: p.pagina === p.totalPaginas ? C.muted : C.slate, opacity: p.pagina === p.totalPaginas ? 0.5 : 1 }}>Siguiente →</button>
      </div>
    </div>
  );
}

function SignalBadge({ estado }) {
  const cfg = ESTADOS[estado] || ESTADOS.pendiente;
  return (
    <div className="inline-flex items-center gap-2">
      <div className="flex items-end gap-0.5" style={{ height: 14 }}>
        {[1, 2, 3].map((i) => (
          <div key={i} style={{ width: 4, height: 4 + i * 3, borderRadius: 1, backgroundColor: i <= cfg.bars ? cfg.color : C.border }} />
        ))}
      </div>
      <span className="text-xs font-medium" style={{ color: cfg.color, fontFamily: 'Inter, sans-serif' }}>{cfg.label}</span>
    </div>
  );
}

function PayPill({ ok, textOk, textPending, onClick }) {
  return (
    <button onClick={onClick} className="text-xs font-semibold px-2.5 py-1 rounded-full transition-colors"
      style={{ backgroundColor: ok ? '#E4F8F1' : '#FFF4E0', color: ok ? C.cyanDark : C.amber, fontFamily: 'Inter, sans-serif', border: `1px solid ${ok ? C.cyanDark : C.amber}33` }}>
      {ok ? textOk : textPending}
    </button>
  );
}

function KPICard({ label, value, sub, accent }) {
  return (
    <div className="rounded-xl p-4" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
      <div className="text-xs uppercase tracking-wide font-semibold" style={{ color: C.muted, fontFamily: 'Inter, sans-serif' }}>{label}</div>
      <div className="text-2xl mt-1.5" style={{ color: accent || C.slate, fontFamily: 'JetBrains Mono, monospace', fontWeight: 600 }}>{value}</div>
      {sub && <div className="text-xs mt-1" style={{ color: C.muted, fontFamily: 'Inter, sans-serif' }}>{sub}</div>}
    </div>
  );
}

function Field({ label, children }) {
  return (
    <label className="flex flex-col gap-1 text-xs" style={{ color: C.muted, fontFamily: 'Inter, sans-serif' }}>
      {label}
      {children}
    </label>
  );
}

const inputStyle = {
  border: `1px solid ${C.border}`, borderRadius: 8, padding: '7px 10px', fontSize: 13,
  fontFamily: 'Inter, sans-serif', color: C.slate, outline: 'none', backgroundColor: '#FAFBFC',
};

export default function App() {
  const [tab, setTab] = useState('dashboard');
  // ---------- Usuarios, sesión y roles ----------
  const [usuarios, setUsuarios] = useState([]);
  const [empresas, setEmpresas] = useState([]);
  const [empresaActiva, setEmpresaActiva] = useState(null);
  const [showEmpresaForm, setShowEmpresaForm] = useState(false);
  const [empresaMenuOpen, setEmpresaMenuOpen] = useState(false);
  const [sesion, setSesion] = useState(null);
  const [authLoading, setAuthLoading] = useState(true);
  const [authMode, setAuthMode] = useState('login'); // 'login' | 'signup'
  MONEDA_SIMBOLO = sesion?.preferencias?.moneda || '$';
  // t('texto en español') -> traduce si el usuario eligió italiano; si no hay traducción registrada, muestra el español tal cual.
  const t = (es) => (sesion?.preferencias?.idioma === 'it' ? (I18N_IT[es] || es) : es);
  const [loginEmail, setLoginEmail] = useState('');
  const [loginClave, setLoginClave] = useState('');
  const [loginNombre, setLoginNombre] = useState('');
  const [loginError, setLoginError] = useState('');
  const [showUsuarioForm, setShowUsuarioForm] = useState(false);
  const [facturasManuales, setFacturasManuales] = useState([]);
  const [showFacturaManualForm, setShowFacturaManualForm] = useState(false);
  const [userMenuOpen, setUserMenuOpen] = useState(false);
  const [cuentaClaveNueva, setCuentaClaveNueva] = useState('');
  const [cuentaClaveConfirm, setCuentaClaveConfirm] = useState('');
  const [cuentaNombreEdit, setCuentaNombreEdit] = useState('');
  const [cuentaUsernameEdit, setCuentaUsernameEdit] = useState('');
  const [cuentaMsg, setCuentaMsg] = useState('');
  const [emailTemplates, setEmailTemplates] = useState(SEED_EMAIL_TEMPLATES); // plantillas: locales por ahora (fase futura)
  const [plantillaSel, setPlantillaSel] = useState('agenda_squadra');
  const [notificaciones, setNotificaciones] = useState([]);
  const [notifPanelOpen, setNotifPanelOpen] = useState(false);
  const [confirmState, setConfirmState] = useState(null); // { mensaje, onConfirm } | null
  // Confirmación propia del sistema (no depende de window.confirm, que algunos visores bloquean en silencio)
  const pedirConfirmacion = (mensaje, onConfirm) => setConfirmState({ mensaje, onConfirm });
  const [rolesConfig, setRolesConfig] = useState(DEFAULT_ROLES_CONFIG);
  const [rolesLectura, setRolesLectura] = useState({}); // rol -> [claves en solo lectura]
  const [rolesOcultarValores, setRolesOcultarValores] = useState({ Tecnico: ['ordenes', 'agenda'] }); // rol -> [módulos donde no ve montos $]
  const [rolesPorNombre, setRolesPorNombre] = useState({});
  const [miembrosEmpresa, setMiembrosEmpresa] = useState([]); // personas con acceso real a la empresa activa
  const [invitacionesEmpresa, setInvitacionesEmpresa] = useState([]); // invitaciones pendientes
  const [inviteEmail, setInviteEmail] = useState('');
  const [inviteRol, setInviteRol] = useState(''); // nombre de rol -> UUID real en Supabase
  const [nuevoRolNombre, setNuevoRolNombre] = useState('');
  const rolesLista = Object.keys(rolesConfig);
  const tieneAcceso = (rol, key) => {
    const permisos = rolesConfig[rol];
    if (permisos === null || permisos === undefined) return rol === 'Administrador';
    return permisos.includes(key) || (rolesLectura[rol] || []).includes(key);
  };
  const esSoloLectura = (rol, key) => rol !== 'Administrador' && (rolesLectura[rol] || []).includes(key) && !(rolesConfig[rol] || []).includes(key);
  // Punto 6: si el rol tiene ocultos los valores monetarios en este módulo, los montos se muestran como "•••••"
  const ocultaValores = (key) => sesion && sesion.rol !== 'Administrador' && (rolesOcultarValores[sesion.rol] || []).includes(key);
  const fmtSeg = (monto, key) => ocultaValores(key || tab) ? '•••••' : fmt(monto);
  // Guardia para acciones de escritura: true = bloqueado (avisa y no ejecuta)
  const bloqueadoPorLectura = () => {
    if (sesion && esSoloLectura(sesion.rol, tab)) { alert('Tu rol tiene acceso de solo lectura en este módulo.'); return true; }
    return false;
  };
  const setPermisoNivel = (rol, key, nivel) => {
    const arrEditar = (rolesConfig[rol] || []).filter((k) => k !== key);
    const nuevoEditar = nivel === 'editar' ? [...arrEditar, key] : arrEditar;
    const arrLectura = (rolesLectura[rol] || []).filter((k) => k !== key);
    const nuevoLectura = nivel === 'ver' ? [...arrLectura, key] : arrLectura;
    setRolesConfig((prev) => ({ ...prev, [rol]: nuevoEditar }));
    setRolesLectura((prev) => ({ ...prev, [rol]: nuevoLectura }));
    const rolId = rolesPorNombre[rol];
    if (rolId) supabase.from('roles').update({ permisos: nuevoEditar, permisos_lectura: nuevoLectura }).eq('id', rolId).then(({ error }) => { if (error) console.error(error); });
  };
  const toggleOcultarValores = (rol, key) => {
    const arr = rolesOcultarValores[rol] || [];
    const nuevo = arr.includes(key) ? arr.filter((k) => k !== key) : [...arr, key];
    setRolesOcultarValores((prev) => ({ ...prev, [rol]: nuevo }));
    const rolId = rolesPorNombre[rol];
    if (rolId) supabase.from('roles').update({ ocultar_valores: nuevo }).eq('id', rolId).then(({ error }) => { if (error) console.error(error); });
  };
  const addRol = async (nombre) => {
    if (rolesConfig[nombre]) { alert('Ese rol ya existe.'); return; }
    const { data: nuevo, error } = await supabase.from('roles').insert({ empresa_id: empresaActiva, nombre, es_administrador: false, permisos: [], permisos_lectura: [], ocultar_valores: [] }).select().single();
    if (error) { alert('No se pudo crear el rol: ' + error.message); return; }
    setRolesConfig((prev) => ({ ...prev, [nombre]: [] }));
    setRolesLectura((prev) => ({ ...prev, [nombre]: [] }));
    setRolesOcultarValores((prev) => ({ ...prev, [nombre]: [] }));
    setRolesPorNombre((prev) => ({ ...prev, [nombre]: nuevo.id }));
  };
  const eliminarRol = (nombre) => {
    const rolId = rolesPorNombre[nombre];
    setRolesConfig((prev) => { const n = { ...prev }; delete n[nombre]; return n; });
    if (rolId) supabase.from('roles').delete().eq('id', rolId).then(({ error }) => { if (error) console.error(error); });
  };
  const invitarUsuario = async () => {
    const email = inviteEmail.trim().toLowerCase();
    if (!email || !inviteRol) { alert('Escribe un correo y elige un rol.'); return; }
    if (miembrosEmpresa.some((m) => m.email === email)) { alert('Esa persona ya tiene acceso a esta empresa.'); return; }
    const rolId = rolesPorNombre[inviteRol];
    const { error } = await supabase.from('invitaciones').insert({ empresa_id: empresaActiva, email, rol_id: rolId, creada_por: sesion?.id || null });
    if (error) { alert(error.message.includes('duplicate') ? 'Ya existe una invitación pendiente para ese correo.' : 'No se pudo invitar: ' + error.message); return; }
    setInvitacionesEmpresa((prev) => [...prev, { id: crypto.randomUUID?.() || String(Date.now()), email, rol: inviteRol, rolId }]);
    setInviteEmail('');
    alert(`Invitación registrada para ${email}. No se envía ningún correo automático todavía — avísale tú mismo que entre al sitio y cree su cuenta (o inicie sesión, si ya la tiene) con ese mismo correo exacto; en ese momento obtendrá acceso automáticamente.`);
  };
  const cancelarInvitacion = (id) => {
    setInvitacionesEmpresa((prev) => prev.filter((i) => i.id !== id));
    supabase.from('invitaciones').delete().eq('id', id).then(({ error }) => { if (error) console.error(error); });
  };
  const cambiarRolMiembro = (perfilId, nuevoRolNombreSel) => {
    const rolId = rolesPorNombre[nuevoRolNombreSel];
    if (!rolId) return;
    setMiembrosEmpresa((prev) => prev.map((m) => m.perfilId === perfilId ? { ...m, rol: nuevoRolNombreSel, rolId } : m));
    supabase.from('perfil_empresas').update({ rol_id: rolId }).eq('perfil_id', perfilId).eq('empresa_id', empresaActiva).then(({ error }) => { if (error) console.error(error); });
  };
  // Vincula el perfil de un usuario con rol Técnico a su registro de empleado — sin esto, el filtro
  // "solo ve sus propias órdenes" no tiene con qué comparar y no restringe nada.
  const vincularEmpleadoMiembro = (perfilId, empleadoId) => {
    setMiembrosEmpresa((prev) => prev.map((m) => m.perfilId === perfilId ? { ...m, empleadoId } : m));
    supabase.from('perfiles').update({ empleado_id: empleadoId }).eq('id', perfilId).then(({ error }) => { if (error) console.error(error); });
    if (perfilId === sesion?.id) setSesion((prev) => ({ ...prev, empleadoId }));
  };
  const cambiarNombreMiembro = (perfilId, nuevoNombre) => {
    setMiembrosEmpresa((prev) => prev.map((m) => m.perfilId === perfilId ? { ...m, nombre: nuevoNombre } : m));
    supabase.from('perfiles').update({ nombre: nuevoNombre }).eq('id', perfilId).then(({ error }) => { if (error) console.error('Error al cambiar nombre:', error.message); });
    if (perfilId === sesion?.id) setSesion((prev) => ({ ...prev, nombre: nuevoNombre }));
  };
  // Designa (o quita) a una persona la capacidad de ver avisos financieros sensibles (saldos de bancos).
  const toggleAvisosFinancieros = (perfilId, valor) => {
    setMiembrosEmpresa((prev) => prev.map((m) => m.perfilId === perfilId ? { ...m, recibeAvisosFinancieros: valor } : m));
    supabase.from('perfil_empresas').update({ recibe_avisos_financieros: valor }).eq('perfil_id', perfilId).eq('empresa_id', empresaActiva)
      .then(({ error }) => { if (error) console.error('Error al designar avisos financieros:', error.message); });
  };
  const quitarAccesoMiembro = (perfilId) => {
    if (perfilId === sesion?.id) { alert('No puedes quitarte tu propio acceso desde aquí.'); return; }
    pedirConfirmacion('¿Quitar el acceso de esta persona a la empresa? Podrás volver a invitarla después.', () => {
      setMiembrosEmpresa((prev) => prev.filter((m) => m.perfilId !== perfilId));
      supabase.from('perfil_empresas').delete().eq('perfil_id', perfilId).eq('empresa_id', empresaActiva).then(({ error }) => { if (error) console.error(error); });
    });
  };

  // Recuerda si el usuario prefiere ver su Agenda en celular por semana o por mes.
  useEffect(() => {
    if (sesion?.preferencias?.vistaAgenda) setVistaAgendaMovil(sesion.preferencias.vistaAgenda);
  }, [sesion?.id]);

  // Carga las notificaciones propias del usuario en sesión (las más recientes primero).
  useEffect(() => {
    if (!sesion?.id) { setNotificaciones([]); return; }
    supabase.from('notificaciones').select('*').eq('usuario_id', sesion.id).order('creada_en', { ascending: false }).limit(50).then(({ data }) => {
      setNotificaciones((data || []).map((n) => ({ id: n.id, usuarioId: n.usuario_id, tipo: n.tipo, titulo: n.titulo, mensaje: n.mensaje, ordenId: null, fecha: n.creada_en, leida: n.leida })));
    });

    // Notificaciones en vivo: en cuanto llega una nueva a mi nombre, aparece sola en la campana — sin recargar la página.
    const canal = supabase
      .channel('notificaciones-' + sesion.id)
      .on('postgres_changes', { event: 'INSERT', schema: 'public', table: 'notificaciones', filter: `usuario_id=eq.${sesion.id}` }, (payload) => {
        const n = payload.new;
        setNotificaciones((prev) => {
          if (prev.some((x) => x.id === n.id)) return prev; // evita duplicar si ya llegó por otra vía
          return [{ id: n.id, usuarioId: n.usuario_id, tipo: n.tipo, titulo: n.titulo, mensaje: n.mensaje, ordenId: null, fecha: n.creada_en, leida: n.leida }, ...prev].slice(0, 50);
        });
      })
      .subscribe();
    return () => { supabase.removeChannel(canal); };
  }, [sesion?.id]);

  // Al iniciar sesión, si había invitaciones pendientes hechas ANTES de que esta persona ya tuviera cuenta
  // (o llegadas después de su registro), se reclaman automáticamente aquí.
  useEffect(() => {
    if (!sesion?.email) return;
    supabase.from('invitaciones').select('id, empresa_id, rol_id').eq('email', sesion.email).eq('aceptada', false).then(async ({ data }) => {
      if (!data || data.length === 0) return;
      for (const inv of data) {
        await supabase.from('perfil_empresas').insert({ perfil_id: sesion.id, empresa_id: inv.empresa_id, rol_id: inv.rol_id }).then(() => {});
        await supabase.from('invitaciones').update({ aceptada: true }).eq('id', inv.id);
      }
      // Refresca la lista de empresas del usuario para que la nueva aparezca de inmediato.
      supabase.from('empresas').select('*').then(({ data: emps }) => setEmpresas((emps || []).map((e) => ({ id: e.id, nombre: e.nombre, giro: e.giro, eslogan: e.eslogan, activa: e.activa }))));
    });
  }, [sesion?.email]);
  const [clients, setClients] = useState([]);
  const [materials, setMaterials] = useState([]);
  const [orders, setOrders] = useState([]);
  const [projects, setProjects] = useState([]);
  const [kardexLog, setKardexLog] = useState([]);

  const [showOrderForm, setShowOrderForm] = useState(false);
  const [importPreview, setImportPreview] = useState([]);
  const [agendaBase, setAgendaBase] = useState(new Date().toISOString().slice(0, 10));
  const [agendaFiltro, setAgendaFiltro] = useState('todas');
  const [agendaSel, setAgendaSel] = useState(null);
  const [vistaAgendaMovil, setVistaAgendaMovil] = useState('semana'); // 'semana' | 'mes' — se sincroniza con la preferencia guardada al iniciar sesión
  const [mesAgendaBase, setMesAgendaBase] = useState(new Date().toISOString().slice(0, 10));
  const [diaMesSel, setDiaMesSel] = useState(new Date().toISOString().slice(0, 10));
  const [agendaEdit, setAgendaEdit] = useState(false);
  const [agendaDraft, setAgendaDraft] = useState({});
  const [erModo, setErModo] = useState('periodo');
  const [erMostrarDetalle, setErMostrarDetalle] = useState(true);
  const [compareFrom, setCompareFrom] = useState('2026-06-01');
  const [compareTo, setCompareTo] = useState('2026-06-30');
  const [cierreFecha, setCierreFecha] = useState('');
  const [prorrateoMetodo, setProrrateoMetodo] = useState('ventas');
  const [prorrateoMes, setProrrateoMes] = useState('2026-07');
  const [prorrateosGuardados, setProrrateosGuardados] = useState([]);
  const [bonoUmbral, setBonoUmbral] = useState(10);
  const [bonoMonto, setBonoMonto] = useState(1000);
  const [bonoMes, setBonoMes] = useState('2026-07');
  const [nominaMes, setNominaMes] = useState('2026-07');
  const [showClientForm, setShowClientForm] = useState(false);
  const [showMatForm, setShowMatForm] = useState(false);
  const [showProjectForm, setShowProjectForm] = useState(false);
  const [orderFilter, setOrderFilter] = useState('todas');
  const [ordersSearch, setOrdersSearch] = useState('');
  const [kardexMaterial, setKardexMaterial] = useState('');
  const [kardexAction, setKardexAction] = useState(null);
  // Si el material seleccionado en Kardex ya no existe (o todavía no se ha elegido ninguno), se selecciona el primero real.
  useEffect(() => {
    if (materials.length > 0 && !materials.some((m) => m.id === kardexMaterial)) setKardexMaterial(materials[0].id);
  }, [materials]);
  const [movimientos, setMovimientos] = useState([]);
  const [showMovForm, setShowMovForm] = useState(false);
  const [movFilter, setMovFilter] = useState('todos');
  const [reportFrom, setReportFrom] = useState('2026-07-01');
  const [reportTo, setReportTo] = useState('2026-07-31');
  const [reportType, setReportType] = useState('estado_resultados');
  const [proveedores, setProveedores] = useState([]);
  const [cxp, setCxp] = useState([]);
  const [showProvForm, setShowProvForm] = useState(false);
  const [showCxpForm, setShowCxpForm] = useState(false);
  const [activos, setActivos] = useState([]);
  const [showActivoForm, setShowActivoForm] = useState(false);
  const [depreciaciones, setDepreciaciones] = useState([]);
  const [depMes, setDepMes] = useState('2026-07');
  const [empleados, setEmpleados] = useState([]);
  const techs = empleados.filter((e) => e.rol === 'Técnico' && e.activo !== false); // Técnicos activos (los inactivos conservan su historial)
  const [nomina, setNomina] = useState([]);
  const [anticipos, setAnticipos] = useState([]);
  const [showAnticipoForm, setShowAnticipoForm] = useState(false);
  const [prestamos, setPrestamos] = useState([]);
  const [showPrestamoForm, setShowPrestamoForm] = useState(false);
  const [nominaCiclo, setNominaCiclo] = useState('mensual'); // 'mensual' | 'quincena1' | 'quincena2'
  const [showEmpleadoForm, setShowEmpleadoForm] = useState(false);
  const [showNominaForm, setShowNominaForm] = useState(false);
  const [asientos, setAsientos] = useState([]);
  const asientosRef = useRef([]); // evita colisiones de ID cuando postAsiento se llama varias veces seguidas
  const [showAsientoForm, setShowAsientoForm] = useState(false);
  const [editingAsiento, setEditingAsiento] = useState(null);
  const [cuentas, setCuentas] = useState([]);
  const [centros, setCentros] = useState([]);
  const [gastoCategorias, setGastoCategorias] = useState([]);
  const [catalogoPartidas, setCatalogoPartidas] = useState([]);
  const [tiposServicio, setTiposServicio] = useState([]);
  const [nuevoTipoServicio, setNuevoTipoServicio] = useState('');
  const [ingresoCategorias, setIngresoCategorias] = useState([]);
  const [showCentroForm, setShowCentroForm] = useState(false);
  const [configSeccion, setConfigSeccion] = useState('centros');
  const [conciliados, setConciliados] = useState({});
  const [bancoMovs, setBancoMovs] = useState([]);
  const [conciliacionesGuardadas, setConciliacionesGuardadas] = useState([]);
  const [ventas, setVentas] = useState([]);
  const [showVentaForm, setShowVentaForm] = useState(false);
  const [equiposComodato, setEquiposComodato] = useState([]);
  const [showEquipoForm, setShowEquipoForm] = useState(false);
  const [squadras, setSquadras] = useState([]);
  const [showSquadraForm, setShowSquadraForm] = useState(false);
  const [flotaLog, setFlotaLog] = useState([]); // combustible, mantenimiento, multas
  const [showFlotaForm, setShowFlotaForm] = useState(null); // 'combustible' | 'mantenimiento' | 'multa' | null
  const [flotaActivoSel, setFlotaActivoSel] = useState(null);

  // ---------- Multiempresa conectado a Supabase ----------
  // Carga las empresas a las que el usuario en sesión tiene acceso (RLS ya filtra esto en la base de datos)
  useEffect(() => {
    if (!sesion) { setEmpresas([]); return; }
    supabase.from('empresas').select('*').order('creada_en').then(({ data, error }) => {
      if (error) { console.error(error); return; }
      const mapeadas = (data || []).map((e) => ({ id: e.id, nombre: e.nombre, giro: e.giro, eslogan: e.eslogan, activa: e.activa }));
      setEmpresas(mapeadas);
      if (mapeadas.length > 0 && !mapeadas.some((e) => e.id === empresaActiva)) setEmpresaActiva(mapeadas[0].id);
    });
  }, [sesion]);

  // Carga TODOS los módulos de la empresa activa desde Supabase (se dispara al entrar o cambiar de empresa).
  useEffect(() => {
    if (!empresaActiva) return;
    const eq = (tabla) => supabase.from(tabla).select('*').eq('empresa_id', empresaActiva);

    (async () => {
      // 1) Cuentas y centros primero: el resto de módulos referencia sus CÓDIGOS (texto), no el UUID.
      const { data: ctasData } = await eq('cuentas');
      const cuentasCargadas = (ctasData || []).map((c) => ({ id: c.id, codigo: c.codigo, nombre: c.nombre, tipo: c.tipo, subtipo: c.subtipo, grupo: c.grupo }));
      setCuentas(cuentasCargadas);
      const codPorId = Object.fromEntries(cuentasCargadas.map((c) => [c.id, c.codigo]));

      const { data: ctrosData } = await eq('centros_costo');
      const centrosCargados = (ctrosData || []).map((c) => ({ id: c.id, codigo: c.codigo, nombre: c.nombre }));
      setCentros(centrosCargados);

      // Mapa id de perfil -> nombre, para mostrar "Edgar Farias" en vez del UUID en usuario/capturado por/validado por.
      const { data: miembrosNombres } = await supabase.from('perfil_empresas').select('perfil_id, perfiles(nombre)').eq('empresa_id', empresaActiva);
      const nombrePorPerfilId = Object.fromEntries((miembrosNombres || []).map((m) => [m.perfil_id, m.perfiles?.nombre || '—']));

      // 2) Categorías: se guarda el código de cuenta (texto), como espera el resto del sistema.
      const { data: catsData } = await eq('categorias');
      setGastoCategorias((catsData || []).filter((c) => c.tipo === 'gasto').map((c) => ({ nombre: c.nombre, cuenta: codPorId[c.cuenta_id] || '' })));
      setIngresoCategorias((catsData || []).filter((c) => c.tipo === 'ingreso').map((c) => ({ nombre: c.nombre, cuenta: codPorId[c.cuenta_id] || '' })));

      eq('catalogo_partidas').then(({ data }) => setCatalogoPartidas((data || []).map((p) => ({ id: p.id, nombre: p.nombre, precioRef: Number(p.precio_ref) }))));
      eq('tipos_servicio').then(({ data }) => setTiposServicio((data || []).map((t) => t.nombre)));

      // Roles reales de esta empresa: se reconstruyen los mismos objetos que ya usa el resto del sistema.
      const { data: rolesData } = await eq('roles');
      const nuevoRolesConfig = {}, nuevoRolesLectura = {}, nuevoRolesOcultar = {}, nuevoRolesPorNombre = {};
      (rolesData || []).forEach((r) => {
        nuevoRolesConfig[r.nombre] = r.es_administrador ? null : (r.permisos || []);
        nuevoRolesLectura[r.nombre] = r.permisos_lectura || [];
        nuevoRolesOcultar[r.nombre] = r.ocultar_valores || [];
        nuevoRolesPorNombre[r.nombre] = r.id;
      });
      setRolesConfig(nuevoRolesConfig);
      setRolesLectura(nuevoRolesLectura);
      setRolesOcultarValores(nuevoRolesOcultar);
      setRolesPorNombre(nuevoRolesPorNombre);

      // Miembros reales de la empresa (para la página Usuarios y roles) e invitaciones pendientes.
      const { data: miembrosData } = await supabase.from('perfil_empresas').select('perfil_id, rol_id, recibe_avisos_financieros, perfiles(nombre, email, ultimo_acceso, empleado_id), roles(nombre)').eq('empresa_id', empresaActiva);
      setMiembrosEmpresa((miembrosData || []).map((m) => ({ perfilId: m.perfil_id, nombre: m.perfiles?.nombre, email: m.perfiles?.email, ultimoAcceso: m.perfiles?.ultimo_acceso, empleadoId: m.perfiles?.empleado_id, rol: m.roles?.nombre, rolId: m.rol_id, recibeAvisosFinancieros: m.recibe_avisos_financieros })));
      const { data: invsData } = await supabase.from('invitaciones').select('id, email, rol_id, aceptada, roles(nombre)').eq('empresa_id', empresaActiva).eq('aceptada', false);
      setInvitacionesEmpresa((invsData || []).map((i) => ({ id: i.id, email: i.email, rol: i.roles?.nombre, rolId: i.rol_id })));
      supabase.from('cierre_contable').select('*').eq('empresa_id', empresaActiva).maybeSingle().then(({ data }) => setCierreFecha(data?.cierre_fecha || ''));

      eq('empleados').then(({ data }) => setEmpleados((data || []).map((e) => ({
        id: e.id, nombre: e.nombre, rol: e.rol, puesto: e.puesto, telefono: e.telefono,
        fechaIngreso: e.fecha_ingreso, salarioMensual: Number(e.salario_mensual), reportaA: e.reporta_a, activo: e.activo,
      }))));
      eq('clientes').then(({ data }) => setClients((data || []).map((c) => ({ id: c.id, nombre: c.nombre, direccion: c.direccion, telefono: c.telefono, activo: c.activo }))));
      eq('proveedores').then(({ data }) => setProveedores((data || []).map((p) => ({ id: p.id, nombre: p.nombre, contacto: p.contacto, telefono: p.telefono, activo: p.activo }))));
      eq('materiales').then(({ data }) => setMaterials((data || []).map((m) => ({ id: m.id, nombre: m.nombre, unidad: m.unidad, costoUnitario: Number(m.costo_unitario), stock: Number(m.stock) }))));
      eq('activos_fijos').then(({ data }) => setActivos((data || []).map((a) => ({
        id: a.id, nombre: a.nombre, categoria: a.categoria, fechaCompra: a.fecha_compra, costo: Number(a.costo),
        valorResidual: Number(a.valor_residual), vidaUtilAnios: a.vida_util_anios, responsable: a.responsable_id,
        centroCostoId: a.centro_costo_id, placa: a.placa, km: a.km, seguroVencimiento: a.seguro_vencimiento, revisionVencimiento: a.revision_vencimiento,
      }))));
      eq('depreciaciones').then(({ data }) => setDepreciaciones((data || []).map((d) => ({ id: d.id, activoId: d.activo_id, periodo: d.periodo, monto: Number(d.monto), fecha: d.fecha }))));
      eq('proyectos').then(({ data }) => setProjects((data || []).map((p) => ({ id: p.id, nombre: p.nombre, clienteId: p.cliente_id, fechaInicio: p.fecha_inicio, presupuesto: Number(p.presupuesto), estado: p.estado }))));

      eq('ordenes').then(({ data }) => setOrders((data || []).map((o) => ({
        id: o.folio, supabaseId: o.id, clienteId: o.cliente_id, tecnicoId: o.tecnico_id, proyectoId: o.proyecto_id,
        centroCostoId: o.centro_costo_id, tipoServicio: o.tipo_servicio, direccion: o.direccion, fecha: o.fecha, hora: o.hora,
        estado: o.estado, precioCliente: Number(o.precio_cliente), costoManoObra: Number(o.costo_mano_obra),
        estadoPago: o.estado_pago, estadoPagoTecnico: o.estado_pago_tecnico, ingresoPosteado: o.ingreso_posteado,
        validada: o.validada, validadaPor: o.validada_por ? (nombrePorPerfilId[o.validada_por] || '—') : null, fechaValidacion: o.fecha_validacion, capturadaPor: o.capturada_por ? (nombrePorPerfilId[o.capturada_por] || '—') : null,
        referencia: o.referencia, notas: o.notas, squadra: o.squadra, idRisorsa: o.id_risorsa, building: o.building,
        cnoBuilding: o.cno_building, roeBuilding: o.roe_building, materialesUsados: o.materiales_usados || [],
        partidas: o.partidas || [], reprogramaciones: o.reprogramaciones || [],
      }))));
      eq('kardex').then(({ data }) => setKardexLog((data || []).map((k) => ({
        id: k.id, materialId: k.material_id, fecha: k.fecha, tipo: k.tipo, cantidad: Number(k.cantidad),
        costoUnitario: Number(k.costo_unitario), referencia: k.referencia, saldoCantidad: Number(k.saldo_cantidad),
        saldoValor: Number(k.saldo_valor), centroCostoId: k.centro_costo_id,
      }))));
      eq('squadras').then(({ data }) => setSquadras((data || []).map((s) => ({ id: s.id, nombre: s.nombre, miembros: s.miembros || [], activa: s.activa }))));
      eq('equipos_comodato').then(({ data }) => setEquiposComodato((data || []).map((e) => ({
        id: e.id, tipo: e.tipo, modelo: e.modelo, serial: e.serial, sector: e.sector, estado: e.estado,
        clienteId: e.cliente_id, ordenId: e.orden_id, fechaAsignacion: e.fecha_asignacion, fechaDevolucion: e.fecha_devolucion,
      }))));
      eq('flota_log').then(({ data }) => setFlotaLog((data || []).map((f) => ({
        id: f.id, tipo: f.tipo, activoId: f.activo_id, fecha: f.fecha, monto: Number(f.monto),
        kmActual: f.km_actual, subtipo: f.subtipo, detalle: f.detalle,
      }))));

      eq('ventas').then(({ data }) => setVentas((data || []).map((v) => ({
        id: v.folio, supabaseId: v.id, clienteId: v.cliente_id, fecha: v.fecha, lineas: v.lineas || [],
        total: Number(v.total), costo: Number(v.costo), estadoPago: v.estado_pago, anulada: v.anulada, usuario: v.usuario_id,
      }))));
      eq('cxp').then(({ data }) => setCxp((data || []).map((c) => ({
        id: c.id, proveedorId: c.proveedor_id, concepto: c.concepto, tipo: c.tipo, categoria: c.categoria,
        materialId: c.material_id, cantidad: c.cantidad, fecha: c.fecha, vencimiento: c.vencimiento, monto: Number(c.monto), estado: c.estado,
      }))));
      eq('facturas_manuales').then(({ data }) => setFacturasManuales((data || []).map((f) => ({
        id: f.id, clienteId: f.cliente_id, concepto: f.concepto, categoria: f.categoria, fecha: f.fecha,
        monto: Number(f.monto), cuentaIngreso: codPorId[f.cuenta_ingreso_id] || '', estadoPago: f.estado_pago,
      }))));
      eq('movimientos').then(({ data }) => setMovimientos((data || []).map((m) => ({
        id: m.id, tipo: m.tipo, categoria: m.categoria, descripcion: m.descripcion, fecha: m.fecha, monto: Number(m.monto),
      }))));

      eq('nomina').then(({ data }) => setNomina((data || []).map((n) => ({
        id: n.id, empleadoId: n.empleado_id, periodo: n.periodo, fechaPago: n.fecha_pago, montoBruto: Number(n.monto_bruto),
        descuentoAnticipo: Number(n.descuento_anticipo), descuentoPrestamo: Number(n.descuento_prestamo), monto: Number(n.monto), estado: n.estado,
      }))));
      eq('anticipos').then(({ data }) => setAnticipos((data || []).map((a) => ({ id: a.id, empleadoId: a.empleado_id, fecha: a.fecha, monto: Number(a.monto), motivo: a.motivo, estado: a.estado }))));
      eq('prestamos').then(({ data }) => setPrestamos((data || []).map((p) => ({
        id: p.id, empleadoId: p.empleado_id, fecha: p.fecha, motivo: p.motivo, montoTotal: Number(p.monto_total),
        cuotas: p.cuotas, montoCuota: Number(p.monto_cuota), saldoPendiente: Number(p.saldo_pendiente), cuotasPagadas: p.cuotas_pagadas, estado: p.estado,
      }))));

      // 3) Asientos: se resuelven a código de cuenta usando el mapa ya calculado arriba (sin fetch redundante).
      const { data: asData } = await eq('asientos');
      const mapeados = (asData || []).map((a) => ({
        id: 'AS' + String(a.id).replace(/-/g, '').slice(0, 6).toUpperCase(), supabaseId: a.id, fecha: a.fecha, glosa: a.glosa,
        cuentaDebe: codPorId[a.cuenta_debe_id] || a.cuenta_debe_id, montoDebe: Number(a.monto_debe),
        cuentaHaber: codPorId[a.cuenta_haber_id] || a.cuenta_haber_id, montoHaber: Number(a.monto_haber),
        centroCostoId: a.centro_costo_id, anulado: a.anulado, usuario: a.usuario_id ? (nombrePorPerfilId[a.usuario_id] || '—') : 'Sistema',
      }));
      asientosRef.current = mapeados;
      setAsientos(mapeados);
    })();
  }, [empresaActiva]);

  const cambiarEmpresa = (nuevaId) => {
    if (nuevaId === empresaActiva) { setEmpresaMenuOpen(false); return; }
    setEmpresaActiva(nuevaId);
    setEmpresaMenuOpen(false);
    setUserMenuOpen(false);
    setTab(sesion?.rol === 'Tecnico' ? 'agenda' : 'dashboard');
  };
  const empresasDeSesion = () => empresas.filter((e) => e.activa !== false);

  // Crea la empresa en Supabase, el rol Administrador de esa empresa, y conecta a quien la crea como su admin.
  const addEmpresa = async (data) => {
    const { data: nueva, error } = await supabase.from('empresas')
      .insert({ nombre: data.nombre, giro: data.giro, eslogan: data.eslogan || null }).select().single();
    if (error) { alert('No se pudo crear la empresa: ' + error.message); return; }
    const { data: rolAdmin, error: errRol } = await supabase.from('roles')
      .insert({ empresa_id: nueva.id, nombre: 'Administrador', es_administrador: true }).select().single();
    if (errRol) { alert('Empresa creada, pero falló crear su rol Administrador: ' + errRol.message); return; }
    await supabase.from('perfil_empresas').insert({ perfil_id: sesion.id, empresa_id: nueva.id, rol_id: rolAdmin.id });
    // Semilla de catálogos genéricos para que la empresa nueva no arranque totalmente vacía
    const cuentasBase = SEED_PLAN_CUENTAS.map((c) => ({ empresa_id: nueva.id, codigo: c.codigo, nombre: c.nombre, tipo: c.tipo, subtipo: c.subtipo || null, grupo: c.grupo || null }));
    await supabase.from('cuentas').insert(cuentasBase);
    await supabase.from('centros_costo').insert([{ empresa_id: nueva.id, codigo: 'CC-OPE', nombre: 'Operaciones' }, { empresa_id: nueva.id, codigo: 'CC-ADM', nombre: 'Administrativo' }]);
    await supabase.from('tipos_servicio').insert({ empresa_id: nueva.id, nombre: 'Servicio general' });
    setEmpresas((prev) => [...prev, { id: nueva.id, nombre: nueva.nombre, giro: nueva.giro, eslogan: nueva.eslogan, activa: true }]);
  };
  const empresaActual = empresas.find((e) => e.id === empresaActiva) || empresas[0];
  const [saldoBanco, setSaldoBanco] = useState('');
  const [conciliacionCorte, setConciliacionCorte] = useState(new Date().toISOString().slice(0, 10));
  const [showCatGastoForm, setShowCatGastoForm] = useState(false);
  const [showCatIngresoForm, setShowCatIngresoForm] = useState(false);
  const [showCuentaForm, setShowCuentaForm] = useState(false);
  const [cuentaMayor, setCuentaMayor] = useState('');
  const [mayorCentro, setMayorCentro] = useState('todos');
  const [contabVista, setContabVista] = useState('estado_resultados');
  const [expandedGroup, setExpandedGroup] = useState(null);
  const [mobileNavOpen, setMobileNavOpen] = useState(false);
  const [estadoResCentro, setEstadoResCentro] = useState('todos');

  const clientName = (id) => clients.find((c) => c.id === id)?.nombre || '—';
  const techName = (id) => empleados.find((t) => t.id === id)?.nombre || '—';
  const projectName = (id) => projects.find((p) => p.id === id)?.nombre || '—';
  const centroCostoName = (id) => centros.find((c) => c.id === id)?.nombre || 'Sin centro';
  const matById = (id) => materials.find((m) => m.id === id);

  const materialCost = (order) =>
    order.materialesUsados.reduce((sum, m) => {
      const mat = matById(m.materialId);
      return sum + (mat ? mat.costoUnitario * m.cantidad : 0);
    }, 0);
  const orderCost = (order) => materialCost(order) + Number(order.costoManoObra || 0);
  const orderProfit = (order) => Number(order.precioCliente || 0) - orderCost(order);

  const activeOrders = useMemo(() => orders.filter((o) => o.estado !== 'cancelada'), [orders]);
  // Órdenes con ingreso reconocido en Cuentas por Cobrar: todas las que ya pasaron validación (no canceladas, no por validar)
  const completedOrders = useMemo(() => orders.filter((o) => o.ingresoPosteado), [orders]);
  const receivable = useMemo(() => completedOrders.filter((o) => o.estadoPago === 'pendiente').reduce((s, o) => s + Number(o.precioCliente || 0), 0), [completedOrders]);
  const techPending = useMemo(() => activeOrders.filter((o) => o.estadoPagoTecnico === 'pendiente').reduce((s, o) => s + Number(o.costoManoObra || 0), 0), [activeOrders]);
  const inProgressCount = orders.filter((o) => o.estado === 'pendiente' || o.estado === 'en_proceso').length;
  const lowStock = materials.filter((m) => m.stock < 10);
  const inventoryValue = materials.reduce((s, m) => s + m.stock * m.costoUnitario, 0);
  const activeProjects = projects.filter((p) => p.estado === 'activo').length;

  // Registra un asiento contable automático (partida simple: debe = haber).
  const postAsiento = (fecha, glosa, cuentaDebe, monto, cuentaHaber, centroCostoId = null) => {
    if (!monto) return;
    // Control de cierre: los asientos automáticos no pueden caer en un periodo cerrado;
    // se registran con fecha actual y nota, para no alterar meses ya conciliados.
    let fechaFinal = fecha, glosaFinal = glosa;
    if (cierreFecha && fecha <= cierreFecha) {
      fechaFinal = new Date().toISOString().slice(0, 10);
      glosaFinal = `${glosa} (op. de periodo cerrado ${fecha})`;
    }
    const idLocal = 'AS' + String(asientosRef.current.length + 1).padStart(3, '0');
    const nuevoLocal = {
      id: idLocal, fecha: fechaFinal, glosa: glosaFinal,
      cuentaDebe, montoDebe: monto, cuentaHaber, montoHaber: monto, centroCostoId,
      usuario: sesion?.nombre || 'Sistema',
    };
    asientosRef.current = [...asientosRef.current, nuevoLocal];
    setAsientos(asientosRef.current);
    // Persistencia real en Supabase: se resuelven los códigos (texto) a sus UUID reales.
    if (empresaActiva) {
      const cD = cuentas.find((c) => c.codigo === cuentaDebe)?.id;
      const cH = cuentas.find((c) => c.codigo === cuentaHaber)?.id;
      const ctro = centroCostoId ? centros.find((c) => c.id === centroCostoId || c.codigo === centroCostoId)?.id : null;
      if (cD && cH) {
        supabase.from('asientos').insert({
          empresa_id: empresaActiva, fecha: fechaFinal, glosa: glosaFinal,
          cuenta_debe_id: cD, monto_debe: monto, cuenta_haber_id: cH, monto_haber: monto,
          centro_costo_id: ctro || null, usuario_id: sesion?.id || null,
        }).then(({ error }) => { if (error) console.error('Error al guardar asiento en Supabase:', error.message); });
      } else {
        console.error('No se pudo resolver la cuenta contable al guardar en Supabase:', cuentaDebe, cuentaHaber);
      }
    }
  };
  // Punto 8: saldo actual de Bancos (1110) para avisos de fondos
  const saldoBancosActual = () => asientos.filter((a) => !a.anulado).reduce((s, a) => s + (a.cuentaDebe === '1110' ? a.montoDebe : 0) - (a.cuentaHaber === '1110' ? a.montoHaber : 0), 0);
  // Aviso de fondos: NUNCA bloquea la acción (evita botones "muertos" si el navegador restringe diálogos nativos).
  // Si el pago dejaría Bancos en negativo, se registra una notificación visible en la campana para Administración.
  const avisoFondos = (monto, concepto) => {
    const saldo = saldoBancosActual();
    if (Number(monto) > saldo) {
      const destinatarios = usuariosAvisosFinancieros();
      if (destinatarios.length > 0) notificar(destinatarios, 'aviso_fondos', 'Aviso de fondos', `El pago de ${concepto} (${fmt(Number(monto))}) dejó Bancos en ${fmt(saldo - Number(monto))}.`);
    }
  };

  // Crea una orden y registra automáticamente la salida de materiales en el kardex.
  // ---------- Importación de órdenes desde Excel (agenda de activaciones) ----------
  const importarOrdenesExcel = (file) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'array', cellDates: true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
        const pad = (n) => String(n).padStart(2, '0');
        const parsed = rows.map((r, i) => {
          const get = (needles) => {
            for (const k of Object.keys(r)) { if (needles.some((t) => k.toLowerCase().includes(t))) return r[k]; }
            return '';
          };
          const cita = get(['appuntamento', 'cita', 'fecha', 'date']);
          let fecha = '', hora = '';
          const d = cita instanceof Date ? cita : (cita ? new Date(cita) : null);
          if (d && !isNaN(d)) {
            fecha = `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
            hora = `${pad(d.getHours())}:${pad(d.getMinutes())}`;
          }
          const direccion = [get(['indirizzo', 'direccion', 'dirección', 'address']), get(['località', 'localita', 'ciudad', 'localidad', 'city'])].filter(Boolean).join(', ');
          return {
            idx: i, incluir: true,
            referencia: String(get(['descrizione olo', 'olo', 'referencia', 'codigo', 'código']) || ''),
            wr: String(get(['wr']) || ''),
            notas: String(get(['dis.', 'dis', 'nota']) || ''),
            squadra: String(get(['squadra', 'equipo', 'team']) || ''),
            idRisorsa: String(get(['id risorsa', 'risorsa', 'recurso']) || ''),
            building: String((() => { for (const k of Object.keys(r)) { if (k.toLowerCase().trim() === 'building') return r[k]; } return ''; })() || ''),
            cnoBuilding: String(get(['cno building', 'cno']) || ''),
            roeBuilding: String(get(['roe building', 'roe']) || ''),
            direccion,
            fecha: fecha || new Date().toISOString().slice(0, 10),
            hora,
            tecnicoId: techs[0]?.id || '',
            clienteId: clients[0]?.id || '',
            centroCostoId: centros[0]?.id || '',
            costoManoObra: '', precioCliente: '',
          };
        }).filter((p) => p.referencia || p.direccion);
        if (parsed.length === 0) { alert('No se encontraron órdenes en el archivo. Verifica que tenga columnas como Descrizione OLO / Appuntamento / Indirizzo.'); return; }
        // Detectar duplicados: contra órdenes ya registradas (por referencia OLO o WR) y dentro del mismo archivo
        const vistos = new Set();
        const conDuplicados = parsed.map((p) => {
          const existente = orders.find((o) => o.referencia && ((p.referencia && o.referencia.includes(p.referencia)) || (p.wr && o.referencia.includes(`WR ${p.wr}`))));
          const claveInterna = p.referencia || `${p.direccion}|${p.fecha}`;
          const repetidaEnArchivo = vistos.has(claveInterna);
          vistos.add(claveInterna);
          if (existente) return { ...p, incluir: false, duplicado: `Ya importada como ${existente.id}` };
          if (repetidaEnArchivo) return { ...p, incluir: false, duplicado: 'Repetida en este archivo' };
          return p;
        });
        setImportPreview(conDuplicados);
      } catch (err) {
        alert('No se pudo leer el archivo: ' + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const updateImportRow = (idx, campo, valor) => setImportPreview((prev) => prev.map((p) => p.idx === idx ? { ...p, [campo]: valor } : p));

  const crearOrdenesImportadas = async () => {
    const seleccionadas = importPreview.filter((p) => p.incluir);
    if (seleccionadas.length === 0) return;
    const requiereValidacion = sesion && !['Administrador', 'Contabilidad'].includes(sesion.rol);
    const estadoInicial = requiereValidacion ? 'por_validar' : 'pendiente';
    let creadas = 0;
    let listaLocal = [...orders];
    const conteoPorTecnico = {}; // tecnicoId -> cantidad de órdenes que le tocaron, para avisarle una sola vez al final
    for (const p of seleccionadas) {
      const folio = nextFolio(listaLocal);
      const costoManoObra = Number(p.costoManoObra) || 0;
      const precioCliente = Number(p.precioCliente) || 0;
      const referencia = [p.referencia, p.wr && `WR ${p.wr}`].filter(Boolean).join(' · ');

      const { data: nueva, error } = await supabase.from('ordenes').insert({
        empresa_id: empresaActiva, folio, cliente_id: p.clienteId || null, tecnico_id: p.tecnicoId || null,
        centro_costo_id: p.centroCostoId || null, tipo_servicio: 'Activación FTTH', direccion: p.direccion,
        fecha: p.fecha, hora: p.hora || null, estado: estadoInicial, precio_cliente: precioCliente, costo_mano_obra: costoManoObra,
        estado_pago: 'pendiente', estado_pago_tecnico: 'pendiente', ingreso_posteado: false, validada: !requiereValidacion,
        capturada_por: sesion?.id || null, referencia: referencia || null, notas: p.notas || null,
        squadra: p.squadra || null, id_risorsa: p.idRisorsa || null, building: p.building || null,
        cno_building: p.cnoBuilding || null, roe_building: p.roeBuilding || null, materiales_usados: [], partidas: [],
      }).select().single();
      if (error) { console.error('Error al importar orden:', error.message); continue; }

      const ordenLocal = {
        id: folio, supabaseId: nueva.id, clienteId: p.clienteId, tecnicoId: p.tecnicoId, proyectoId: null,
        centroCostoId: p.centroCostoId, tipoServicio: 'Activación FTTH', direccion: p.direccion, referencia,
        notas: p.notas, squadra: p.squadra || null, idRisorsa: p.idRisorsa || null, building: p.building || null,
        cnoBuilding: p.cnoBuilding || null, roeBuilding: p.roeBuilding || null, fecha: p.fecha, hora: p.hora || null,
        estado: estadoInicial, validada: !requiereValidacion, capturadaPor: sesion?.nombre || 'Sistema',
        materialesUsados: [], partidas: [], costoManoObra, precioCliente, estadoPago: 'pendiente', estadoPagoTecnico: 'pendiente',
        ingresoPosteado: false,
      };
      listaLocal = [...listaLocal, ordenLocal];
      creadas++;
      if (estadoInicial === 'pendiente' && p.tecnicoId) conteoPorTecnico[p.tecnicoId] = (conteoPorTecnico[p.tecnicoId] || 0) + 1;

      if (costoManoObra > 0) postAsiento(p.fecha, `Mano de obra por pagar — orden ${folio}`, '6140', costoManoObra, '2510', p.centroCostoId);
      if (estadoInicial === 'pendiente' && precioCliente > 0) {
        postAsiento(p.fecha, `Cuenta por cobrar — orden ${folio}`, '1305', precioCliente, '4135', p.centroCostoId);
        ordenLocal.ingresoPosteado = true;
        supabase.from('ordenes').update({ ingreso_posteado: true }).eq('id', nueva.id).then(() => {});
      }
    }
    setOrders(listaLocal);
    setImportPreview([]);
    // Aviso a cada técnico (una sola notificación consolidada, no una por cada orden importada)
    Object.entries(conteoPorTecnico).forEach(([tecnicoId, cantidad]) => {
      notificar(usuariosPorEmpleado(tecnicoId), 'orden_asignada', 'Nuevas órdenes asignadas',
        `Se te asignaron ${cantidad} orden(es) de trabajo nueva(s) por una importación desde Excel.`, null);
    });
    if (requiereValidacion && creadas > 0) {
      notificar(usuariosAdmin(), 'orden_por_validar', 'Órdenes por validar',
        `${sesion?.nombre || 'Un operador'} importó ${creadas} orden(es) desde Excel que esperan tu validación.`, null);
    }
    alert(`Se crearon ${creadas} de ${seleccionadas.length} orden(es) de trabajo.${creadas < seleccionadas.length ? ' Revisa la consola para ver los errores de las que fallaron.' : ''}`);
  };

  // ---------- Notificaciones in-app (simuladas en esta sesión; con Supabase se disparan aunque el destinatario no tenga la app abierta) ----------
  const usuariosPorEmpleado = (empleadoId) => miembrosEmpresa.filter((m) => m.empleadoId === empleadoId).map((m) => ({ id: m.perfilId }));
  const usuariosAdmin = () => miembrosEmpresa.filter((m) => ['Administrador', 'Contabilidad'].includes(m.rol)).map((m) => ({ id: m.perfilId }));
  // Avisos de fondos/bancos: información sensible — solo la ven las personas que el Administrador haya designado explícitamente.
  const usuariosAvisosFinancieros = () => miembrosEmpresa.filter((m) => m.recibeAvisosFinancieros).map((m) => ({ id: m.perfilId }));
  const notificar = async (destinatarios, tipo, titulo, mensaje, ordenId) => {
    if (!destinatarios || destinatarios.length === 0 || !empresaActiva) return;
    const ordenSupabaseId = ordenId ? orders.find((o) => o.id === ordenId)?.supabaseId || null : null;
    const filas = destinatarios.map((u) => ({ empresa_id: empresaActiva, usuario_id: u.id, tipo, titulo, mensaje, orden_id: ordenSupabaseId }));
    // No se pide .select() de vuelta: cada quien solo puede LEER sus propias notificaciones (por privacidad),
    // así que si aquí se manda una notificación a otra persona, pedir leerla de regreso rompería el guardado completo.
    const { error } = await supabase.from('notificaciones').insert(filas);
    if (error) { console.error('Error al guardar notificación:', error.message); return; }
    // Si una de las notificaciones es para uno mismo, se refleja de inmediato en la campana sin esperar a recargar.
    if (destinatarios.some((u) => u.id === sesion?.id)) {
      const ahora = new Date().toISOString();
      setNotificaciones((prev) => {
        const nuevas = destinatarios.filter((u) => u.id === sesion.id).map((u, i) => ({ id: 'temp-' + Date.now() + i, usuarioId: u.id, tipo, titulo, mensaje, ordenId, fecha: ahora, leida: false }));
        return [...nuevas, ...prev].slice(0, 50);
      });
    }
  };

  const addOrderWithKardex = async (data) => {
    if (bloqueadoPorLectura()) return;
    // Control de stock: no permitir consumir más material del disponible
    const faltantes = (data.materialesUsados || []).map((mu) => {
      const m = materials.find((x) => x.id === mu.materialId);
      return m && mu.cantidad > m.stock ? `${m.nombre}: pides ${mu.cantidad}, hay ${m.stock}` : null;
    }).filter(Boolean);
    if (faltantes.length > 0) { alert('Stock insuficiente:\n' + faltantes.join('\n') + '\n\nRegistra una entrada de material en Kardex antes de crear la orden.'); return; }
    const folio = nextFolio(orders);
    const requiereValidacion = sesion && !['Administrador', 'Contabilidad'].includes(sesion.rol);
    const estadoInicial = requiereValidacion ? 'por_validar' : 'pendiente';

    const { data: nueva, error } = await supabase.from('ordenes').insert({
      empresa_id: empresaActiva, folio, cliente_id: data.clienteId || null, tecnico_id: data.tecnicoId || null,
      proyecto_id: data.proyectoId || null, centro_costo_id: data.centroCostoId || null, tipo_servicio: data.tipoServicio,
      direccion: data.direccion, fecha: data.fecha, hora: data.hora || null, estado: estadoInicial,
      precio_cliente: Number(data.precioCliente) || 0, costo_mano_obra: Number(data.costoManoObra) || 0,
      estado_pago: 'pendiente', estado_pago_tecnico: 'pendiente', ingreso_posteado: false, validada: !requiereValidacion,
      capturada_por: sesion?.id || null, referencia: data.referencia || null, notas: data.notas || null,
      squadra: data.squadra || null, id_risorsa: data.idRisorsa || null, building: data.building || null,
      cno_building: data.cnoBuilding || null, roe_building: data.roeBuilding || null,
      materiales_usados: data.materialesUsados || [], partidas: data.partidas || [],
    }).select().single();
    if (error) { alert('No se pudo guardar la orden: ' + error.message); return; }

    // Descontar stock (local + Supabase) y registrar kardex de salida por cada material usado
    const matSnapshot = materials.map((m) => ({ ...m }));
    const newEntries = [];
    for (const mu of (data.materialesUsados || [])) {
      const idx = matSnapshot.findIndex((m) => m.id === mu.materialId);
      if (idx === -1) continue;
      matSnapshot[idx].stock -= mu.cantidad;
      const { data: kNuevo } = await supabase.from('kardex').insert({
        empresa_id: empresaActiva, material_id: mu.materialId, fecha: data.fecha, tipo: 'salida', cantidad: mu.cantidad,
        costo_unitario: matSnapshot[idx].costoUnitario, referencia: `Orden ${folio}`,
        saldo_cantidad: matSnapshot[idx].stock, saldo_valor: matSnapshot[idx].stock * matSnapshot[idx].costoUnitario,
        centro_costo_id: data.centroCostoId || null,
      }).select().single();
      await supabase.from('materiales').update({ stock: matSnapshot[idx].stock }).eq('id', mu.materialId);
      if (kNuevo) newEntries.push({
        id: kNuevo.id, materialId: mu.materialId, fecha: data.fecha, tipo: 'salida', cantidad: mu.cantidad,
        costoUnitario: matSnapshot[idx].costoUnitario, referencia: `Orden ${folio}`,
        saldoCantidad: matSnapshot[idx].stock, saldoValor: matSnapshot[idx].stock * matSnapshot[idx].costoUnitario, centroCostoId: data.centroCostoId || null,
      });
    }
    setMaterials(matSnapshot);
    setKardexLog((prev) => [...prev, ...newEntries]);

    const ordenLocal = {
      id: folio, supabaseId: nueva.id, estadoPago: 'pendiente', estadoPagoTecnico: 'pendiente', ...data,
      estado: estadoInicial, validada: !requiereValidacion, capturadaPor: sesion?.nombre || 'Sistema', ingresoPosteado: false,
    };
    setOrders((prev) => [...prev, ordenLocal]);

    const costoMaterialesOrden = newEntries.reduce((s, k) => s + k.cantidad * k.costoUnitario, 0);
    if (costoMaterialesOrden > 0) postAsiento(data.fecha, `Consumo de materiales — orden ${folio}`, '6135', costoMaterialesOrden, '1435', data.centroCostoId);
    if (Number(data.costoManoObra) > 0) postAsiento(data.fecha, `Mano de obra por pagar — orden ${folio}`, '6140', Number(data.costoManoObra), '2510', data.centroCostoId);
    // El ingreso se reconoce de inmediato en Cuentas por Cobrar (no se espera a que la orden se complete);
    // si requiere validación, se reconoce hasta que administración la valide (ver validarOrden).
    if (estadoInicial === 'pendiente' && Number(data.precioCliente) > 0) {
      postAsiento(data.fecha, `Cuenta por cobrar — orden ${folio}`, '1305', Number(data.precioCliente), '4135', data.centroCostoId);
      setOrders((prev) => prev.map((o) => o.id === folio ? { ...o, ingresoPosteado: true } : o));
      supabase.from('ordenes').update({ ingreso_posteado: true }).eq('id', nueva.id).then(() => {});
    }
    // Notificaciones: si ya quedó activa, avisa al técnico asignado; si requiere validación, avisa a administración.
    if (estadoInicial === 'pendiente' && data.tecnicoId) {
      notificar(usuariosPorEmpleado(data.tecnicoId), 'orden_asignada', 'Nueva orden asignada', `Se te asignó la orden ${folio}${data.fecha ? ` para el ${data.fecha}` : ''}.`, folio);
    } else if (estadoInicial === 'por_validar') {
      notificar(usuariosAdmin(), 'orden_por_validar', 'Orden por validar', `La orden ${folio}, capturada por ${sesion?.nombre || 'un operador'}, espera tu validación.`, folio);
    }
  };

  // Validación de órdenes: administración revisa lo cargado por el operador antes de que entre a la agenda.
  // Al validar, se reconoce el ingreso en Cuentas por Cobrar (si no venía ya reconocido).
  const validarOrden = (id, aprobar) => {
    if (bloqueadoPorLectura()) return;
    const ord = orders.find((o) => o.id === id);
    if (!ord || ord.estado !== 'por_validar') return;
    const ingresoPosteadoAhora = aprobar && Number(ord.precioCliente) > 0 && !ord.ingresoPosteado;
    if (ingresoPosteadoAhora) {
      postAsiento(ord.fecha, `Cuenta por cobrar — orden ${id} (validada)`, '1305', Number(ord.precioCliente), '4135', ord.centroCostoId);
    }
    setOrders((prev) => prev.map((o) => {
      if (o.id !== id || o.estado !== 'por_validar') return o;
      return aprobar
        ? { ...o, estado: 'pendiente', validada: true, validadaPor: sesion?.nombre, fechaValidacion: new Date().toISOString().slice(0, 10), ingresoPosteado: Number(o.precioCliente) > 0 }
        : { ...o, estado: 'cancelada', validada: false, validadaPor: sesion?.nombre, fechaValidacion: new Date().toISOString().slice(0, 10), notas: `${o.notas || ''} [Rechazada en validación]`.trim() };
    }));
    if (ord.supabaseId) {
      supabase.from('ordenes').update({
        estado: aprobar ? 'pendiente' : 'cancelada', validada: !!aprobar, validada_por: sesion?.id || null,
        fecha_validacion: new Date().toISOString().slice(0, 10), ingreso_posteado: ingresoPosteadoAhora || ord.ingresoPosteado,
        notas: aprobar ? ord.notas : `${ord.notas || ''} [Rechazada en validación]`.trim(),
      }).eq('id', ord.supabaseId).then(({ error }) => { if (error) console.error(error); });
    }
    if (aprobar && ord.tecnicoId) {
      notificar(usuariosPorEmpleado(ord.tecnicoId), 'orden_asignada', 'Nueva orden asignada', `Se validó y te fue asignada la orden ${id}${ord.fecha ? ` para el ${ord.fecha}` : ''}.`, id);
    }
  };

  // Cambia el estado operativo de una orden (agenda/ejecución). El ingreso ya se reconoció al crear/validar la orden,
  // así que completar/en proceso NO vuelve a tocar Cuentas por Cobrar — solo cancelar revierte lo ya reconocido.
  const changeOrderEstado = async (id, nuevoEstado) => {
    if (bloqueadoPorLectura()) return;
    const ordActual = orders.find((o) => o.id === id);
    // Control: no cancelar si la MO ya se pagó al técnico (primero revertir el pago)
    if (nuevoEstado === 'cancelada' && ordActual && ordActual.estadoPagoTecnico === 'pagado') {
      alert('La mano de obra de esta orden ya fue pagada al técnico. Revierte ese pago antes de cancelar.');
      return;
    }
    if (nuevoEstado === 'cancelada' && ordActual && ordActual.estadoPago === 'pagada') {
      alert('Esta orden ya fue cobrada. Primero desmarca el cobro en Facturación antes de cancelarla.');
      return;
    }
    // Devolución de materiales al cancelar: regresan al inventario y se revierte el costo
    if (nuevoEstado === 'cancelada' && ordActual && ordActual.estado !== 'cancelada') {
      const salidas = kardexLog.filter((k) => k.tipo === 'salida' && k.referencia === `Orden ${id}`);
      const yaDevueltas = kardexLog.some((k) => k.tipo === 'entrada' && k.referencia === `Devolución orden ${id}`);
      if (salidas.length > 0 && !yaDevueltas) {
        const matActualizados = materials.map((m) => ({ ...m }));
        const entradas = [];
        let costoDevuelto = 0;
        for (const s of salidas) {
          const idx = matActualizados.findIndex((m) => m.id === s.materialId);
          if (idx === -1) continue;
          matActualizados[idx].stock += s.cantidad;
          costoDevuelto += s.cantidad * s.costoUnitario;
          const { data: kNuevo } = await supabase.from('kardex').insert({
            empresa_id: empresaActiva, material_id: s.materialId, fecha: new Date().toISOString().slice(0, 10), tipo: 'entrada',
            cantidad: s.cantidad, costo_unitario: s.costoUnitario, referencia: `Devolución orden ${id}`,
            saldo_cantidad: matActualizados[idx].stock, saldo_valor: matActualizados[idx].stock * matActualizados[idx].costoUnitario,
            centro_costo_id: ordActual.centroCostoId || null,
          }).select().single();
          await supabase.from('materiales').update({ stock: matActualizados[idx].stock }).eq('id', s.materialId);
          if (kNuevo) entradas.push({
            id: kNuevo.id, materialId: s.materialId, fecha: new Date().toISOString().slice(0, 10), tipo: 'entrada', cantidad: s.cantidad,
            costoUnitario: s.costoUnitario, referencia: `Devolución orden ${id}`,
            saldoCantidad: matActualizados[idx].stock, saldoValor: matActualizados[idx].stock * matActualizados[idx].costoUnitario, centroCostoId: ordActual.centroCostoId || null,
          });
        }
        setMaterials(matActualizados);
        setKardexLog((prev) => [...prev, ...entradas]);
        if (costoDevuelto > 0) postAsiento(new Date().toISOString().slice(0, 10), `Devolución de materiales — orden ${id} cancelada`, '1435', costoDevuelto, '6135', ordActual.centroCostoId);
      }
      // Reversión de la mano de obra comprometida (aún no pagada)
      if (Number(ordActual.costoManoObra) > 0) {
        postAsiento(new Date().toISOString().slice(0, 10), `Reversión MO — orden ${id} cancelada`, '2510', Number(ordActual.costoManoObra), '6140', ordActual.centroCostoId);
      }
      // Reversión de la cuenta por cobrar ya reconocida
      if (ordActual.ingresoPosteado && Number(ordActual.precioCliente) > 0) {
        postAsiento(new Date().toISOString().slice(0, 10), `Reversión CxC — orden ${id} cancelada`, '4135', Number(ordActual.precioCliente), '1305', ordActual.centroCostoId);
      }
    }
    setOrders((prev) => prev.map((o) => {
      if (o.id !== id) return o;
      if (nuevoEstado === 'cancelada') return { ...o, estado: nuevoEstado, ingresoPosteado: false };
      return { ...o, estado: nuevoEstado };
    }));
    if (ordActual?.supabaseId) {
      supabase.from('ordenes').update({ estado: nuevoEstado, ingreso_posteado: nuevoEstado === 'cancelada' ? false : ordActual.ingresoPosteado }).eq('id', ordActual.supabaseId)
        .then(({ error }) => { if (error) console.error(error); });
    }
    // Notifica a administración ante cualquier cambio de estado hecho por alguien que no sea Administrador/Contabilidad
    // (así se enteran aunque el técnico solo mueva a "en proceso" o cancele, no solo cuando completa).
    if (ordActual && ordActual.estado !== nuevoEstado && sesion && !['Administrador', 'Contabilidad'].includes(sesion.rol)) {
      const etiquetas = { pendiente: 'Pendiente', en_proceso: 'En proceso', completada: 'Completada', cancelada: 'Cancelada' };
      const tipo = nuevoEstado === 'completada' ? 'orden_completada' : 'orden_cambio_estado';
      const titulo = nuevoEstado === 'completada' ? 'Orden completada' : `Orden ${etiquetas[nuevoEstado] || nuevoEstado}`;
      notificar(usuariosAdmin(), tipo, titulo,
        `${sesion?.nombre || 'Un técnico'} cambió la orden ${id} a "${etiquetas[nuevoEstado] || nuevoEstado}"${ordActual.tecnicoId ? ` (${empleadoName(ordActual.tecnicoId)})` : ''}.`, id);
    }
  };

  // Marca el pago del cliente y registra el cobro en caja; al desmarcar, revierte.
  const togglePagoCliente = (id) => {
    if (bloqueadoPorLectura()) return;
    const ord = orders.find((o) => o.id === id);
    const nuevo = ord?.estadoPago === 'pagada' ? 'pendiente' : 'pagada';
    setOrders((prev) => prev.map((o) => {
      if (o.id !== id) return o;
      const hoy = new Date().toISOString().slice(0, 10);
      if (nuevo === 'pagada') postAsiento(hoy, `Cobro de cliente — orden ${o.id}`, '1110', Number(o.precioCliente), '1305', o.centroCostoId);
      else postAsiento(hoy, `Reversión cobro de cliente — orden ${o.id}`, '1305', Number(o.precioCliente), '1110', o.centroCostoId);
      return { ...o, estadoPago: nuevo };
    }));
    if (ord?.supabaseId) supabase.from('ordenes').update({ estado_pago: nuevo }).eq('id', ord.supabaseId).then(({ error }) => { if (error) console.error(error); });
  };

  // Marca el pago al técnico y liquida el pasivo de mano de obra.
  const togglePagoTecnico = (id) => {
    if (bloqueadoPorLectura()) return;
    const ord = orders.find((o) => o.id === id);
    if (ord && ord.estadoPagoTecnico !== 'pagado') avisoFondos(ord.costoManoObra, `pago técnico ${id}`);
    const nuevo = ord?.estadoPagoTecnico === 'pagado' ? 'pendiente' : 'pagado';
    setOrders((prev) => prev.map((o) => {
      if (o.id !== id) return o;
      const hoy = new Date().toISOString().slice(0, 10);
      if (nuevo === 'pagado') postAsiento(hoy, `Pago de mano de obra — orden ${o.id}`, '2510', Number(o.costoManoObra), '1110', o.centroCostoId);
      else postAsiento(hoy, `Reversión pago de mano de obra — orden ${o.id}`, '1110', Number(o.costoManoObra), '2510', o.centroCostoId);
      return { ...o, estadoPagoTecnico: nuevo };
    }));
    if (ord?.supabaseId) supabase.from('ordenes').update({ estado_pago_tecnico: nuevo }).eq('id', ord.supabaseId).then(({ error }) => { if (error) console.error(error); });
  };

  // Elimina una orden y revierte al inventario los materiales que había consumido.
  // Solo se permite eliminar una orden que YA está cancelada — así se garantiza que su contabilidad
  // (CxC, mano de obra, materiales) ya fue revertida por changeOrderEstado antes de borrar el registro.
  const deleteOrderWithReversal = (id) => {
    if (bloqueadoPorLectura()) return;
    const order = orders.find((o) => o.id === id);
    if (!order) return;
    if (order.estado !== 'cancelada') {
      alert('Solo se pueden eliminar órdenes ya canceladas — así se garantiza que su contabilidad quedó revertida. Cancélala primero desde el selector de estado.');
      return;
    }
    pedirConfirmacion(`¿Eliminar definitivamente la orden ${id}? Esta acción no se puede deshacer.`, () => {
      setOrders((prev) => prev.filter((o) => o.id !== id));
      if (order.supabaseId) supabase.from('ordenes').delete().eq('id', order.supabaseId).then(({ error }) => { if (error) console.error(error); });
    });
  };

  // Registra una entrada (compra) de material y recalcula el costo promedio ponderado.
  const registerEntrada = async ({ materialId, cantidad, costoUnitario, fecha, referencia, centroCostoId }) => {
    const idx = materials.findIndex((m) => m.id === materialId);
    if (idx === -1) return;
    const m = materials[idx];
    const qty = Number(cantidad);
    const cu = Number(costoUnitario);
    const newStock = m.stock + qty;
    const newCosto = newStock > 0 ? (m.stock * m.costoUnitario + qty * cu) / newStock : cu;
    const updated = [...materials];
    updated[idx] = { ...m, stock: newStock, costoUnitario: newCosto };
    setMaterials(updated);
    supabase.from('materiales').update({ stock: newStock, costo_unitario: newCosto }).eq('id', materialId).then(({ error }) => { if (error) console.error(error); });
    const { data: kNuevo } = await supabase.from('kardex').insert({
      empresa_id: empresaActiva, material_id: materialId, fecha, tipo: 'entrada', cantidad: qty, costo_unitario: cu,
      referencia: referencia || 'Compra', saldo_cantidad: newStock, saldo_valor: newStock * newCosto, centro_costo_id: centroCostoId || null,
    }).select().single();
    if (kNuevo) setKardexLog((prev) => [...prev, {
      id: kNuevo.id, materialId, fecha, tipo: 'entrada', cantidad: qty, costoUnitario: cu,
      referencia: referencia || 'Compra', saldoCantidad: newStock, saldoValor: newStock * newCosto, centroCostoId: centroCostoId || null,
    }]);
    postAsiento(fecha, `Compra de material — ${m.nombre}`, '1435', qty * cu, '1110', centroCostoId || null);
  };

  // Registra un consumo manual de material (salida) NO ligado a una orden, asignado a un centro de costo.
  const registerSalida = async ({ materialId, cantidad, fecha, referencia, centroCostoId, ordenSupabaseId }) => {
    const idx = materials.findIndex((m) => m.id === materialId);
    if (idx === -1) return;
    const m = materials[idx];
    const qty = Math.min(Number(cantidad) || 0, m.stock);
    if (qty <= 0) return;
    const nuevoStock = m.stock - qty;
    const updated = [...materials];
    updated[idx] = { ...m, stock: nuevoStock };
    setMaterials(updated);
    supabase.from('materiales').update({ stock: nuevoStock }).eq('id', materialId).then(({ error }) => { if (error) console.error(error); });
    const { data: kNuevo } = await supabase.from('kardex').insert({
      empresa_id: empresaActiva, material_id: materialId, fecha, tipo: 'salida', cantidad: qty, costo_unitario: m.costoUnitario,
      referencia: referencia || 'Consumo manual', saldo_cantidad: nuevoStock, saldo_valor: nuevoStock * m.costoUnitario, centro_costo_id: centroCostoId || null,
    }).select().single();
    if (kNuevo) setKardexLog((prev) => [...prev, {
      id: kNuevo.id, materialId, fecha, tipo: 'salida', cantidad: qty, costoUnitario: m.costoUnitario,
      referencia: referencia || 'Consumo manual', saldoCantidad: nuevoStock, saldoValor: nuevoStock * m.costoUnitario, centroCostoId: centroCostoId || null,
    }]);
    postAsiento(fecha, `Consumo de material — ${m.nombre}${referencia ? ' (' + referencia + ')' : ''}`, '6135', qty * m.costoUnitario, '1435', centroCostoId || null);
    // Si se ligó a una orden de trabajo, se agrega al detalle de esa orden (para que aparezca en su ficha)
    if (ordenSupabaseId) {
      setOrders((prev) => prev.map((o) => o.supabaseId === ordenSupabaseId
        ? { ...o, materialesUsados: [...(o.materialesUsados || []), { materialId, cantidad: qty }] } : o));
      const ordenActual = orders.find((o) => o.supabaseId === ordenSupabaseId);
      const materialesActualizados = [...(ordenActual?.materialesUsados || []), { materialId, cantidad: qty }];
      supabase.from('ordenes').update({ materiales_usados: materialesActualizados }).eq('id', ordenSupabaseId)
        .then(({ error }) => { if (error) console.error('Error al ligar material a la orden:', error.message); });
    }
  };

  const addMaterialWithKardex = async (data) => {
    const stock = Number(data.stock) || 0;
    const costo = Number(data.costoUnitario) || 0;
    const { data: nuevo, error } = await supabase.from('materiales')
      .insert({ empresa_id: empresaActiva, nombre: data.nombre, unidad: data.unidad, costo_unitario: costo, stock }).select().single();
    if (error) { alert('No se pudo guardar el material: ' + error.message); return; }
    setMaterials((p) => [...p, { id: nuevo.id, nombre: nuevo.nombre, unidad: nuevo.unidad, costoUnitario: costo, stock }]);
    const { data: kNuevo } = await supabase.from('kardex').insert({
      empresa_id: empresaActiva, material_id: nuevo.id, fecha: new Date().toISOString().slice(0, 10),
      tipo: 'entrada', cantidad: stock, costo_unitario: costo, referencia: 'Alta de material — stock inicial',
      saldo_cantidad: stock, saldo_valor: stock * costo,
    }).select().single();
    if (kNuevo) setKardexLog((prev) => [...prev, {
      id: kNuevo.id, materialId: nuevo.id, fecha: kNuevo.fecha, tipo: 'entrada', cantidad: stock, costoUnitario: costo,
      referencia: 'Alta de material — stock inicial', saldoCantidad: stock, saldoValor: stock * costo,
    }]);
    if (stock * costo > 0) postAsiento(new Date().toISOString().slice(0, 10), `Alta de inventario — ${data.nombre}`, '1435', stock * costo, '3105');
  };

  const centroCostoMetrics = (centroId, ordenesBase) => {
    const linked = ordenesBase.filter((o) => o.centroCostoId === centroId);
    const ingresos = linked.reduce((s, o) => s + Number(o.precioCliente || 0), 0);
    const cMateriales = linked.reduce((s, o) => s + materialCost(o), 0);
    const cManoObra = linked.reduce((s, o) => s + Number(o.costoManoObra || 0), 0);
    return { ordenes: linked.length, ingresos, cMateriales, cManoObra, utilidad: ingresos - cMateriales - cManoObra };
  };

  const projectMetrics = (p) => {
    const linked = orders.filter((o) => o.proyectoId === p.id);
    const gastado = linked.filter((o) => o.estado !== 'cancelada').reduce((s, o) => s + orderCost(o), 0);
    const ingresos = linked.filter((o) => o.estado === 'completada').reduce((s, o) => s + Number(o.precioCliente || 0), 0);
    return { linked, gastado, ingresos, utilidad: ingresos - gastado, avance: p.presupuesto ? gastado / p.presupuesto : 0 };
  };

  // ---------- Facturas excepcionales (no ligadas a órdenes; el ingreso normal viene de las órdenes) ----------
  const addFacturaManual = async (data) => {
    const monto = Number(data.monto) || 0;
    const cuentaIngresoCodigo = ingresoCategorias.find((c) => c.nombre === data.categoria)?.cuenta || '4200';
    const cuentaIngresoId = cuentas.find((c) => c.codigo === cuentaIngresoCodigo)?.id;
    const { data: nueva, error } = await supabase.from('facturas_manuales').insert({
      empresa_id: empresaActiva, cliente_id: data.clienteId, concepto: data.concepto, categoria: data.categoria,
      cuenta_ingreso_id: cuentaIngresoId || null, fecha: data.fecha, monto, estado_pago: 'pendiente',
    }).select().single();
    if (error) { alert('No se pudo guardar la factura: ' + error.message); return; }
    setFacturasManuales((p) => [...p, { id: nueva.id, ...data, monto, cuentaIngreso: cuentaIngresoCodigo, estadoPago: 'pendiente' }]);
    postAsiento(data.fecha, `Factura excepcional — ${data.concepto}`, '1305', monto, cuentaIngresoCodigo, 'CC-ADM');
  };
  const toggleFacturaManualPago = (id) => {
    const f = facturasManuales.find((x) => x.id === id);
    if (!f) return;
    const nuevo = f.estadoPago === 'pagada' ? 'pendiente' : 'pagada';
    setFacturasManuales((prev) => prev.map((x) => {
      if (x.id !== id) return x;
      const hoy = new Date().toISOString().slice(0, 10);
      if (nuevo === 'pagada') postAsiento(hoy, `Cobro factura excepcional — ${x.concepto}`, '1110', x.monto, '1305', 'CC-ADM');
      else postAsiento(hoy, `Reversión cobro factura excepcional — ${x.concepto}`, '1305', x.monto, '1110', 'CC-ADM');
      return { ...x, estadoPago: nuevo };
    }));
    supabase.from('facturas_manuales').update({ estado_pago: nuevo }).eq('id', id).then(({ error }) => { if (error) console.error(error); });
  };

  // ---------- Ventas (empresas de giro Comercio): ingreso + descarga de inventario ----------
  const addVenta = async (data) => {
    if (bloqueadoPorLectura()) return;
    // Validar stock de todas las líneas
    const faltantes = data.lineas.map((l) => {
      const m = materials.find((x) => x.id === l.materialId);
      return m && l.cantidad > m.stock ? `${m.nombre}: pides ${l.cantidad}, hay ${m.stock}` : null;
    }).filter(Boolean);
    if (faltantes.length > 0) { alert('Stock insuficiente:\n' + faltantes.join('\n')); return; }
    const folio = nextId('V', ventas);
    const totalVenta = data.lineas.reduce((s, l) => s + l.cantidad * l.precioUnitario, 0);
    const matActualizados = materials.map((m) => ({ ...m }));
    let costoTotal = 0;
    for (const l of data.lineas) {
      const idx = matActualizados.findIndex((m) => m.id === l.materialId);
      if (idx === -1) continue;
      costoTotal += l.cantidad * matActualizados[idx].costoUnitario;
      matActualizados[idx].stock -= l.cantidad;
    }
    const { data: nueva, error } = await supabase.from('ventas').insert({
      empresa_id: empresaActiva, folio, cliente_id: data.clienteId, fecha: data.fecha, lineas: data.lineas,
      total: totalVenta, costo: costoTotal, estado_pago: 'pendiente', anulada: false, usuario_id: sesion?.id || null,
    }).select().single();
    if (error) { alert('No se pudo guardar la venta: ' + error.message); return; }
    const salidas = [];
    for (const l of data.lineas) {
      const idx = matActualizados.findIndex((m) => m.id === l.materialId);
      if (idx === -1) continue;
      const { data: kNuevo } = await supabase.from('kardex').insert({
        empresa_id: empresaActiva, material_id: l.materialId, fecha: data.fecha, tipo: 'salida', cantidad: l.cantidad,
        costo_unitario: matActualizados[idx].costoUnitario, referencia: `Venta ${folio}`,
        saldo_cantidad: matActualizados[idx].stock, saldo_valor: matActualizados[idx].stock * matActualizados[idx].costoUnitario, centro_costo_id: null,
      }).select().single();
      await supabase.from('materiales').update({ stock: matActualizados[idx].stock }).eq('id', l.materialId);
      if (kNuevo) salidas.push({
        id: kNuevo.id, materialId: l.materialId, fecha: data.fecha, tipo: 'salida', cantidad: l.cantidad,
        costoUnitario: matActualizados[idx].costoUnitario, referencia: `Venta ${folio}`,
        saldoCantidad: matActualizados[idx].stock, saldoValor: matActualizados[idx].stock * matActualizados[idx].costoUnitario, centroCostoId: 'CC-OPE',
      });
    }
    setMaterials(matActualizados);
    setKardexLog((prev) => [...prev, ...salidas]);
    setVentas((p) => [...p, { id: folio, supabaseId: nueva.id, clienteId: data.clienteId, fecha: data.fecha, lineas: data.lineas, total: totalVenta, costo: costoTotal, estadoPago: 'pendiente', anulada: false, usuario: sesion?.nombre }]);
    postAsiento(data.fecha, `Venta ${folio}`, '1305', totalVenta, '4135', 'CC-OPE');
    if (costoTotal > 0) postAsiento(data.fecha, `Costo de ventas — ${folio}`, '6135', costoTotal, '1435', 'CC-OPE');
  };
  const toggleVentaCobro = (id) => {
    if (bloqueadoPorLectura()) return;
    const v = ventas.find((x) => x.id === id);
    if (!v || v.anulada) return;
    const nuevo = v.estadoPago === 'pagada' ? 'pendiente' : 'pagada';
    setVentas((prev) => prev.map((x) => {
      if (x.id !== id) return x;
      const hoy = new Date().toISOString().slice(0, 10);
      if (nuevo === 'pagada') postAsiento(hoy, `Cobro venta ${x.id}`, '1110', x.total, '1305', 'CC-OPE');
      else postAsiento(hoy, `Reversión cobro venta ${x.id}`, '1305', x.total, '1110', 'CC-OPE');
      return { ...x, estadoPago: nuevo };
    }));
    if (v.supabaseId) supabase.from('ventas').update({ estado_pago: nuevo }).eq('id', v.supabaseId).then(({ error }) => { if (error) console.error(error); });
  };
  const anularVenta = (id) => {
    if (bloqueadoPorLectura()) return;
    const v = ventas.find((x) => x.id === id);
    if (!v || v.anulada) return;
    if (v.estadoPago === 'pagada') { alert('Esta venta ya fue cobrada. Revierte el cobro antes de anularla.'); return; }
    pedirConfirmacion(`¿Anular la venta ${id}? Los productos regresarán al inventario y se revertirá la contabilidad.`, async () => {
      const salidasVenta = kardexLog.filter((k) => k.tipo === 'salida' && k.referencia === `Venta ${id}`);
      const matActualizados = materials.map((m) => ({ ...m }));
      const entradas = [];
      for (const s of salidasVenta) {
        const idx = matActualizados.findIndex((m) => m.id === s.materialId);
        if (idx === -1) continue;
        matActualizados[idx].stock += s.cantidad;
        const { data: kNuevo } = await supabase.from('kardex').insert({
          empresa_id: empresaActiva, material_id: s.materialId, fecha: new Date().toISOString().slice(0, 10), tipo: 'entrada',
          cantidad: s.cantidad, costo_unitario: s.costoUnitario, referencia: `Anulación venta ${id}`,
          saldo_cantidad: matActualizados[idx].stock, saldo_valor: matActualizados[idx].stock * matActualizados[idx].costoUnitario, centro_costo_id: null,
        }).select().single();
        await supabase.from('materiales').update({ stock: matActualizados[idx].stock }).eq('id', s.materialId);
        if (kNuevo) entradas.push({
          id: kNuevo.id, materialId: s.materialId, fecha: new Date().toISOString().slice(0, 10), tipo: 'entrada', cantidad: s.cantidad,
          costoUnitario: s.costoUnitario, referencia: `Anulación venta ${id}`,
          saldoCantidad: matActualizados[idx].stock, saldoValor: matActualizados[idx].stock * matActualizados[idx].costoUnitario, centroCostoId: 'CC-OPE',
        });
      }
      setMaterials(matActualizados);
      setKardexLog((prev) => [...prev, ...entradas]);
      const hoy = new Date().toISOString().slice(0, 10);
      postAsiento(hoy, `Anulación venta ${id} — reversión ingreso`, '4135', v.total, '1305', 'CC-OPE');
      if (v.costo > 0) postAsiento(hoy, `Anulación venta ${id} — reversión costo`, '1435', v.costo, '6135', 'CC-OPE');
      setVentas((prev) => prev.map((x) => x.id === id ? { ...x, anulada: true } : x));
      if (v.supabaseId) supabase.from('ventas').update({ anulada: true }).eq('id', v.supabaseId).then(({ error }) => { if (error) console.error(error); });
    });
  };

  // ---------- Equipos en comodato (ONT / CPE): no son consumibles, se rastrean por serial y cliente ----------
  const addEquipoComodato = async (data) => {
    if (bloqueadoPorLectura()) return;
    if (equiposComodato.some((e) => e.serial === data.serial)) { alert('Ya existe un equipo con ese número de serie.'); return; }
    const { data: nuevo, error } = await supabase.from('equipos_comodato').insert({
      empresa_id: empresaActiva, tipo: data.tipo, modelo: data.modelo, serial: data.serial, sector: data.sector, estado: 'en_almacen',
    }).select().single();
    if (error) { alert('No se pudo guardar el equipo: ' + error.message); return; }
    setEquiposComodato((p) => [...p, { id: nuevo.id, ...data, estado: 'en_almacen', clienteId: null, fechaAsignacion: null }]);
  };
  const asignarEquipoComodato = (id, clienteId, ordenId) => {
    if (bloqueadoPorLectura()) return;
    const fecha = new Date().toISOString().slice(0, 10);
    setEquiposComodato((prev) => prev.map((e) => e.id === id ? { ...e, estado: 'asignado', clienteId, ordenId: ordenId || null, fechaAsignacion: fecha } : e));
    supabase.from('equipos_comodato').update({ estado: 'asignado', cliente_id: clienteId, orden_id: ordenId || null, fecha_asignacion: fecha }).eq('id', id).then(({ error }) => { if (error) console.error(error); });
  };
  const devolverEquipoComodato = (id) => {
    if (bloqueadoPorLectura()) return;
    const fecha = new Date().toISOString().slice(0, 10);
    setEquiposComodato((prev) => prev.map((e) => e.id === id ? { ...e, estado: 'en_almacen', clienteId: null, ordenId: null, fechaDevolucion: fecha } : e));
    supabase.from('equipos_comodato').update({ estado: 'en_almacen', cliente_id: null, orden_id: null, fecha_devolucion: fecha }).eq('id', id).then(({ error }) => { if (error) console.error(error); });
  };
  const bajaEquipoComodato = (id) => {
    if (bloqueadoPorLectura()) return;
    pedirConfirmacion('¿Dar de baja este equipo (dañado / perdido)? Se conserva en el historial marcado como baja.', () => {
      setEquiposComodato((prev) => prev.map((e) => e.id === id ? { ...e, estado: 'baja' } : e));
      supabase.from('equipos_comodato').update({ estado: 'baja' }).eq('id', id).then(({ error }) => { if (error) console.error(error); });
    });
  };

  // ---------- Squadras: equipos de técnicos para asignar agenda por cuadrilla ----------
  const addSquadra = async (data) => {
    if (bloqueadoPorLectura()) return;
    const { data: nueva, error } = await supabase.from('squadras').insert({ empresa_id: empresaActiva, nombre: data.nombre, miembros: data.miembros, activa: true }).select().single();
    if (error) { alert('No se pudo guardar la squadra: ' + error.message); return; }
    setSquadras((p) => [...p, { id: nueva.id, nombre: data.nombre, miembros: data.miembros, activa: true }]);
  };
  const toggleSquadraActiva = (id) => {
    if (bloqueadoPorLectura()) return;
    const nuevaActiva = squadras.find((s) => s.id === id)?.activa === false;
    setSquadras((prev) => prev.map((s) => s.id === id ? { ...s, activa: s.activa === false } : s));
    supabase.from('squadras').update({ activa: nuevaActiva }).eq('id', id).then(({ error }) => { if (error) console.error(error); });
  };

  // ---------- Flota: combustible, mantenimiento y multas por vehículo (activos categoría Vehículo) ----------
  const addFlotaRegistro = async (tipo, data) => {
    if (bloqueadoPorLectura()) return;
    const activo = activos.find((a) => a.id === data.activoId);
    const monto = Number(data.monto) || 0;
    const { data: nuevo, error } = await supabase.from('flota_log').insert({
      empresa_id: empresaActiva, tipo, activo_id: data.activoId, fecha: data.fecha, monto,
      km_actual: data.kmActual || null, subtipo: data.subtipo || null, detalle: data.detalle || null,
    }).select().single();
    if (error) { alert('No se pudo guardar el registro de flota: ' + error.message); return; }
    setFlotaLog((p) => [...p, { id: nuevo.id, tipo, ...data, monto }]);
    if (monto > 0) {
      const glosa = { combustible: 'Combustible', mantenimiento: `Mantenimiento (${data.subtipo})`, multa: 'Multa de tránsito' }[tipo];
      postAsiento(data.fecha, `${glosa} — ${activo?.nombre || data.activoId}`, '5195', monto, '1110', activo?.centroCostoId || 'CC-ADM');
    }
    if (tipo === 'combustible' && data.kmActual) {
      setActivos((prev) => prev.map((a) => a.id === data.activoId ? { ...a, km: Number(data.kmActual) } : a));
      supabase.from('activos_fijos').update({ km: Number(data.kmActual) }).eq('id', data.activoId).then(({ error }) => { if (error) console.error(error); });
    }
  };

  const addMovimiento = async (data) => {
    if (bloqueadoPorLectura()) return;
    if (data.tipo === 'gasto') avisoFondos(data.monto, data.descripcion);
    const monto = Number(data.monto) || 0;
    const { data: nuevo, error } = await supabase.from('movimientos').insert({
      empresa_id: empresaActiva, tipo: data.tipo, categoria: data.categoria, descripcion: data.descripcion, fecha: data.fecha, monto,
    }).select().single();
    if (error) { alert('No se pudo guardar el movimiento: ' + error.message); return; }
    setMovimientos((p) => [...p, { id: nuevo.id, ...data, monto }]);
    const lista = data.tipo === 'gasto' ? gastoCategorias : ingresoCategorias;
    const cuentaCategoria = lista.find((c) => c.nombre === data.categoria)?.cuenta || (data.tipo === 'gasto' ? '5195' : '4200');
    if (data.tipo === 'gasto') {
      postAsiento(data.fecha, data.descripcion, cuentaCategoria, monto, '1110', 'CC-ADM');
    } else {
      postAsiento(data.fecha, data.descripcion, '1110', monto, cuentaCategoria, 'CC-ADM');
    }
  };
  const deleteMovimiento = (id) => {
    if (bloqueadoPorLectura()) return;
    const m = movimientos.find((x) => x.id === id);
    if (!m) return;
    pedirConfirmacion(`¿Eliminar este movimiento de ${m.tipo === 'gasto' ? 'gasto' : 'ingreso'} (${fmt(m.monto)})? Se revertirá también su registro contable.`, () => {
      const lista = m.tipo === 'gasto' ? gastoCategorias : ingresoCategorias;
      const cuentaCategoria = lista.find((c) => c.nombre === m.categoria)?.cuenta || (m.tipo === 'gasto' ? '5195' : '4200');
      const hoy = new Date().toISOString().slice(0, 10);
      if (m.tipo === 'gasto') postAsiento(hoy, `Reversión — ${m.descripcion}`, '1110', m.monto, cuentaCategoria, 'CC-ADM');
      else postAsiento(hoy, `Reversión — ${m.descripcion}`, cuentaCategoria, m.monto, '1110', 'CC-ADM');
      setMovimientos((prev) => prev.filter((x) => x.id !== id));
      supabase.from('movimientos').delete().eq('id', id).then(({ error }) => { if (error) console.error(error); });
    });
  };

  const ingresosAdicionales = movimientos.filter((m) => m.tipo === 'ingreso').reduce((s, m) => s + m.monto, 0);
  const gastosAdicionales = movimientos.filter((m) => m.tipo === 'gasto').reduce((s, m) => s + m.monto, 0);

  const inRange = (fecha) => (!reportFrom || fecha >= reportFrom) && (!reportTo || fecha <= reportTo);
  const repOrdenes = orders.filter((o) => o.estado === 'completada' && inRange(o.fecha));
  const repIngresosOrdenes = repOrdenes.reduce((s, o) => s + Number(o.precioCliente || 0), 0);
  const repCostoMateriales = repOrdenes.reduce((s, o) => s + materialCost(o), 0);
  const repCostoManoObra = repOrdenes.reduce((s, o) => s + Number(o.costoManoObra || 0), 0);
  const repUtilidadOperativa = repIngresosOrdenes - repCostoMateriales - repCostoManoObra;
  const repMovimientos = movimientos.filter((m) => inRange(m.fecha));
  const repIngresosAd = repMovimientos.filter((m) => m.tipo === 'ingreso').reduce((s, m) => s + m.monto, 0);
  const repGastosAd = repMovimientos.filter((m) => m.tipo === 'gasto').reduce((s, m) => s + m.monto, 0);
  const repUtilidadNeta = repUtilidadOperativa + repIngresosAd - repGastosAd;
  const gastosPorCategoria = gastoCategorias.map((cat) => ({
    categoria: cat.nombre,
    monto: repMovimientos.filter((m) => m.tipo === 'gasto' && m.categoria === cat.nombre).reduce((s, m) => s + m.monto, 0),
  })).filter((c) => c.monto > 0).sort((a, b) => b.monto - a.monto);
  const maxGastoCat = Math.max(1, ...gastosPorCategoria.map((c) => c.monto));

  const downloadCSV = (filename, rows) => {
    const csv = rows.map((r) => r.map((v) => `"${String(v ?? '').replace(/"/g, '""')}"`).join(',')).join('\n');
    const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename; document.body.appendChild(a); a.click();
    document.body.removeChild(a); URL.revokeObjectURL(url);
  };

  const exportReporte = () => {
    let rows = [];
    let filename = `reporte_${reportType}_${reportFrom}_a_${reportTo}.csv`;
    if (reportType === 'estado_resultados') {
      rows = [
        ['Estado de resultados', `${reportFrom || 'inicio'} a ${reportTo || 'hoy'}`],
        ['Ingresos por instalaciones', repIngresosOrdenes],
        ['Costo de materiales', -repCostoMateriales],
        ['Costo de mano de obra', -repCostoManoObra],
        ['Utilidad operativa', repUtilidadOperativa],
        ['Otros ingresos', repIngresosAd],
        ['Gastos generales', -repGastosAd],
        ['Utilidad neta del periodo', repUtilidadNeta],
      ];
    } else if (reportType === 'ventas') {
      rows = [['Folio', 'Cliente', 'Proyecto', 'Fecha', 'Precio', 'Estado de pago']];
      repOrdenes.forEach((o) => rows.push([o.id, clientName(o.clienteId), o.proyectoId ? projectName(o.proyectoId) : '', o.fecha, o.precioCliente, o.estadoPago]));
    } else if (reportType === 'costos') {
      rows = [['Folio', 'Costo materiales', 'Costo mano de obra', 'Costo total', 'Utilidad']];
      repOrdenes.forEach((o) => rows.push([o.id, materialCost(o), o.costoManoObra, orderCost(o), orderProfit(o)]));
    } else if (reportType === 'gastos_categoria') {
      rows = [['Categoría', 'Monto']];
      gastosPorCategoria.forEach((c) => rows.push([c.categoria, c.monto]));
    } else if (reportType === 'cxc') {
      rows = [['Folio', 'Cliente', 'Fecha', 'Monto']];
      completedOrders.filter((o) => o.estadoPago === 'pendiente' && inRange(o.fecha)).forEach((o) => rows.push([o.id, clientName(o.clienteId), o.fecha, o.precioCliente]));
    } else if (reportType === 'cxp') {
      rows = [['Proveedor', 'Concepto', 'Vencimiento', 'Monto', 'Estado']];
      cxp.forEach((c) => rows.push([proveedorName(c.proveedorId), c.concepto, c.vencimiento, c.monto, c.estado]));
    } else if (reportType === 'inventario') {
      rows = [['Material', 'Stock', 'Costo promedio', 'Valor']];
      materials.forEach((m) => rows.push([m.nombre, m.stock, m.costoUnitario, m.stock * m.costoUnitario]));
    } else if (reportType === 'nomina') {
      rows = [['Empleado', 'Periodo', 'Fecha de pago', 'Monto', 'Estado']];
      nomina.filter((n) => inRange(n.fechaPago)).forEach((n) => rows.push([empleadoName(n.empleadoId), n.periodo, n.fechaPago, n.monto, n.estado]));
    } else if (reportType === 'activos') {
      rows = [['Activo', 'Categoría', 'Centro de costo', 'Costo', 'Dep. mensual', 'Dep. contabilizada', 'Valor en libros contable']];
      activos.forEach((a) => { const d = depreciacionInfo(a); const contab = depContabilizada(a.id); rows.push([a.nombre, a.categoria, centroCostoName(a.centroCostoId), a.costo, d.depAnual / 12, contab, a.costo - contab]); });
    } else if (reportType === 'proyectos') {
      rows = [['Proyecto', 'Cliente', 'Presupuesto', 'Gastado', 'Facturado', 'Utilidad']];
      projects.forEach((p) => { const m = projectMetrics(p); rows.push([p.nombre, clientName(p.clienteId), p.presupuesto, m.gastado, m.ingresos, m.utilidad]); });
    } else if (reportType === 'tecnicos') {
      rows = [['Técnico', 'Órdenes completadas', 'Mano de obra generada', 'Pendiente por pagar']];
      techs.forEach((t) => {
        const done = orders.filter((o) => o.tecnicoId === t.id && o.estado === 'completada');
        const ganado = done.reduce((s, o) => s + Number(o.costoManoObra), 0);
        const pendiente = done.filter((o) => o.estadoPagoTecnico === 'pendiente').reduce((s, o) => s + Number(o.costoManoObra), 0);
        rows.push([t.nombre, done.length, ganado, pendiente]);
      });
    } else if (reportType === 'centros_costo') {
      rows = [['Centro de costo', 'Órdenes', 'Ingresos', 'Costo materiales', 'Costo mano de obra', 'Utilidad']];
      centros.forEach((c) => { const m = centroCostoMetrics(c.id, repOrdenes); rows.push([c.nombre, m.ordenes, m.ingresos, m.cMateriales, m.cManoObra, m.utilidad]); });
      rows.push(['Global', repOrdenes.length, repIngresosOrdenes, repCostoMateriales, repCostoManoObra, repUtilidadOperativa]);
    }
    downloadCSV(filename, rows);
  };

  // ---------- Cuentas por pagar ----------
  const proveedorName = (id) => proveedores.find((p) => p.id === id)?.nombre || '—';
  const addProveedor = async (data) => {
    const { data: nuevo, error } = await supabase.from('proveedores').insert({ empresa_id: empresaActiva, nombre: data.nombre, contacto: data.contacto, telefono: data.telefono }).select().single();
    if (error) { alert('No se pudo guardar el proveedor: ' + error.message); return; }
    setProveedores((p) => [...p, { id: nuevo.id, nombre: nuevo.nombre, contacto: nuevo.contacto, telefono: nuevo.telefono, activo: true }]);
  };
  const addCxp = async (data) => {
    if (bloqueadoPorLectura()) return;
    const monto = Number(data.monto) || 0;
    const { data: nuevo, error } = await supabase.from('cxp').insert({
      empresa_id: empresaActiva, proveedor_id: data.proveedorId, concepto: data.concepto, tipo: data.tipo,
      categoria: data.categoria || null, material_id: data.materialId || null, cantidad: data.cantidad || null,
      lineas: data.lineas || null, fecha: data.fecha, vencimiento: data.vencimiento, monto, estado: 'pendiente',
    }).select().single();
    if (error) { alert('No se pudo guardar la cuenta por pagar: ' + error.message); return; }
    if (data.tipo === 'inventario' && Array.isArray(data.lineas) && data.lineas.length > 0) {
      // Compra a crédito de inventario con varios materiales: cada línea alimenta su propio kardex y costo promedio
      const matActualizados = materials.map((m) => ({ ...m }));
      const nuevasEntradas = [];
      for (const linea of data.lineas) {
        const cantidad = Number(linea.cantidad);
        const costoUnitario = Number(linea.monto) / cantidad;
        const idx = matActualizados.findIndex((m) => m.id === linea.materialId);
        if (idx === -1) continue;
        const stockPrevio = matActualizados[idx].stock, costoPrevio = matActualizados[idx].costoUnitario;
        const nuevoStock = stockPrevio + cantidad;
        const nuevoCostoProm = nuevoStock > 0 ? ((stockPrevio * costoPrevio) + (cantidad * costoUnitario)) / nuevoStock : costoUnitario;
        matActualizados[idx] = { ...matActualizados[idx], stock: nuevoStock, costoUnitario: nuevoCostoProm };
        await supabase.from('materiales').update({ stock: nuevoStock, costo_unitario: nuevoCostoProm }).eq('id', linea.materialId);
        const { data: kNuevo } = await supabase.from('kardex').insert({
          empresa_id: empresaActiva, material_id: linea.materialId, fecha: data.fecha, tipo: 'entrada', cantidad, costo_unitario: costoUnitario,
          referencia: `Compra a crédito — ${proveedorName(data.proveedorId)}`, saldo_cantidad: nuevoStock, saldo_valor: nuevoStock * nuevoCostoProm, centro_costo_id: null,
        }).select().single();
        if (kNuevo) nuevasEntradas.push({
          id: kNuevo.id, materialId: linea.materialId, fecha: data.fecha, tipo: 'entrada', cantidad, costoUnitario,
          referencia: `Compra a crédito — ${proveedorName(data.proveedorId)}`, saldoCantidad: nuevoStock, saldoValor: nuevoStock * nuevoCostoProm, centroCostoId: 'CC-ADM',
        });
      }
      setMaterials(matActualizados);
      setKardexLog((prev) => [...prev, ...nuevasEntradas]);
      postAsiento(data.fecha, `Compra de inventario a crédito — ${data.concepto}`, '1435', monto, '2205', 'CC-ADM');
    } else {
      const cuentaGasto = gastoCategorias.find((c) => c.nombre === data.categoria)?.cuenta || '5195';
      postAsiento(data.fecha, `Compra a crédito — ${data.concepto}`, cuentaGasto, monto, '2205', 'CC-ADM');
    }
    setCxp((p) => [...p, { id: nuevo.id, estado: 'pendiente', ...data, monto }]);
  };
  const toggleCxpPago = (id) => {
    if (bloqueadoPorLectura()) return;
    const cx = cxp.find((c) => c.id === id);
    if (cx && cx.estado !== 'pagada') avisoFondos(cx.monto, `pago proveedor ${cx.concepto}`);
    const nuevo = cx?.estado === 'pagada' ? 'pendiente' : 'pagada';
    setCxp((prev) => prev.map((c) => {
      if (c.id !== id) return c;
      const hoy = new Date().toISOString().slice(0, 10);
      if (nuevo === 'pagada') postAsiento(hoy, `Pago a proveedor — ${c.concepto}`, '2205', c.monto, '1110', 'CC-ADM');
      else postAsiento(hoy, `Reversión pago a proveedor — ${c.concepto}`, '1110', c.monto, '2205', 'CC-ADM');
      return { ...c, estado: nuevo };
    }));
    supabase.from('cxp').update({ estado: nuevo }).eq('id', id).then(({ error }) => { if (error) console.error(error); });
  };
  const cxpPendiente = cxp.filter((c) => c.estado === 'pendiente').reduce((s, c) => s + c.monto, 0);

  // ---------- Activos fijos ----------
  const addActivo = async (data) => {
    const costo = Number(data.costo) || 0, valorResidual = Number(data.valorResidual) || 0, vidaUtilAnios = Number(data.vidaUtilAnios) || 1;
    const { data: nuevo, error } = await supabase.from('activos_fijos').insert({
      empresa_id: empresaActiva, nombre: data.nombre, categoria: data.categoria, fecha_compra: data.fechaCompra,
      costo, valor_residual: valorResidual, vida_util_anios: vidaUtilAnios, responsable_id: data.responsable || null,
      centro_costo_id: data.centroCostoId || null, placa: data.placa || null, km: data.km || null,
      seguro_vencimiento: data.seguroVencimiento || null, revision_vencimiento: data.revisionVencimiento || null,
    }).select().single();
    if (error) { alert('No se pudo guardar el activo: ' + error.message); return; }
    setActivos((p) => [...p, {
      id: nuevo.id, nombre: nuevo.nombre, categoria: nuevo.categoria, fechaCompra: nuevo.fecha_compra, costo,
      valorResidual, vidaUtilAnios, responsable: nuevo.responsable_id, centroCostoId: nuevo.centro_costo_id,
      placa: nuevo.placa, km: nuevo.km, seguroVencimiento: nuevo.seguro_vencimiento, revisionVencimiento: nuevo.revision_vencimiento,
    }]);
    postAsiento(data.fechaCompra, `Compra de activo fijo — ${data.nombre}`, '1540', costo, '1110', data.centroCostoId || null);
  };
  // Fecha de referencia fija para toda la sesión (solo cambia de un día para otro, no en cada render)
  const hoyDepreciacion = useMemo(() => new Date(new Date().toISOString().slice(0, 10)), []);
  const depreciacionInfo = (a) => {
    const anios = Math.max(0, (hoyDepreciacion - new Date(a.fechaCompra)) / (1000 * 60 * 60 * 24 * 365));
    const depAnual = (a.costo - a.valorResidual) / a.vidaUtilAnios;
    const depAcumulada = Math.min(depAnual * anios, a.costo - a.valorResidual);
    return { depAnual, depAcumulada, valorEnLibros: a.costo - depAcumulada };
  };
  const valorActivosLibros = useMemo(() => activos.reduce((s, a) => s + depreciacionInfo(a).valorEnLibros, 0), [activos, hoyDepreciacion]);
  // Depreciación ya contabilizada (asientos reales), por activo.
  const depContabilizada = (activoId) => depreciaciones.filter((d) => d.activoId === activoId).reduce((s, d) => s + d.monto, 0);
  // Registra en el libro diario la depreciación mensual de todos los activos, cada uno a SU centro de costo.
  const registrarDepreciacionMes = (periodo) => {
    activos.forEach(async (a) => {
      const yaRegistrado = depreciaciones.some((d) => d.activoId === a.id && d.periodo === periodo);
      if (yaRegistrado) return;
      const depreciable = a.costo - a.valorResidual;
      const depMensual = depreciable / a.vidaUtilAnios / 12;
      const acumuladaPrevia = depContabilizada(a.id);
      const monto = Math.min(depMensual, Math.max(0, depreciable - acumuladaPrevia));
      if (monto <= 0) return;
      // Activos del centro Administrativo deprecian a Gasto (5310); operativos a Costo (6310)
      const cuentaDep = esCentroAdmin(a.centroCostoId) ? '5310' : '6310';
      postAsiento(`${periodo}-01`, `Depreciación ${periodo} — ${a.nombre}`, cuentaDep, monto, '1541', a.centroCostoId || null);
      const { data: nueva } = await supabase.from('depreciaciones').insert({ empresa_id: empresaActiva, activo_id: a.id, periodo, fecha: `${periodo}-01`, monto }).select().single();
      if (nueva) setDepreciaciones((prev) => [...prev, { id: nueva.id, activoId: a.id, periodo, fecha: `${periodo}-01`, monto }]);
    });
  };

  // ---------- Recursos humanos ----------
  const toggleEmpleadoActivo = (id) => {
    if (bloqueadoPorLectura()) return;
    const nuevoValor = !(empleados.find((e) => e.id === id)?.activo === false);
    setEmpleados((prev) => prev.map((e) => e.id === id ? { ...e, activo: !nuevoValor ? true : false } : e));
    supabase.from('empleados').update({ activo: empleados.find((e) => e.id === id)?.activo === false }).eq('id', id)
      .then(({ error }) => { if (error) console.error(error); });
  };
  const updateEmpleado = (id, campo, valor) => {
    const valorFinal = campo.startsWith('salario') || campo.startsWith('tarifa') ? Number(valor) || 0 : valor;
    setEmpleados((prev) => prev.map((e) => e.id === id ? { ...e, [campo]: valorFinal } : e));
    const columnaDb = { salarioMensual: 'salario_mensual', reportaA: 'reporta_a' }[campo] || campo;
    supabase.from('empleados').update({ [columnaDb]: valorFinal }).eq('id', id).then(({ error }) => { if (error) console.error(error); });
  };
  const toggleClienteActivo = (id) => {
    const nuevo = clients.find((c) => c.id === id)?.activo === false;
    setClients((prev) => prev.map((c) => c.id === id ? { ...c, activo: c.activo === false ? true : false } : c));
    supabase.from('clientes').update({ activo: nuevo }).eq('id', id).then(({ error }) => { if (error) console.error(error); });
  };
  const updateCliente = (id, campo, valor) => {
    setClients((prev) => prev.map((c) => c.id === id ? { ...c, [campo]: valor } : c));
    supabase.from('clientes').update({ [campo]: valor }).eq('id', id).then(({ error }) => { if (error) console.error(error); });
  };
  const toggleProveedorActivo = (id) => {
    const nuevo = proveedores.find((p) => p.id === id)?.activo === false;
    setProveedores((prev) => prev.map((p) => p.id === id ? { ...p, activo: p.activo === false ? true : false } : p));
    supabase.from('proveedores').update({ activo: nuevo }).eq('id', id).then(({ error }) => { if (error) console.error(error); });
  };
  const updateProveedor = (id, campo, valor) => {
    setProveedores((prev) => prev.map((p) => p.id === id ? { ...p, [campo]: valor } : p));
    supabase.from('proveedores').update({ [campo]: valor }).eq('id', id).then(({ error }) => { if (error) console.error(error); });
  };
  const addEmpleado = async (data) => {
    const { data: nuevo, error } = await supabase.from('empleados').insert({
      empresa_id: empresaActiva, nombre: data.nombre, rol: data.rol || 'Administrativo', puesto: data.puesto,
      telefono: data.telefono, fecha_ingreso: data.fechaIngreso, salario_mensual: Number(data.salarioMensual) || 0, reporta_a: data.reportaA || null,
    }).select().single();
    if (error) { alert('No se pudo guardar el empleado: ' + error.message); return; }
    setEmpleados((p) => [...p, {
      id: nuevo.id, nombre: nuevo.nombre, rol: nuevo.rol, puesto: nuevo.puesto, telefono: nuevo.telefono,
      fechaIngreso: nuevo.fecha_ingreso, salarioMensual: Number(nuevo.salario_mensual), reportaA: nuevo.reporta_a, activo: true,
    }]);
  };
  const empleadoName = (id) => empleados.find((e) => e.id === id)?.nombre || '—';
  // ---------- Anticipos de sueldo ----------
  const addAnticipo = async (data) => {
    if (bloqueadoPorLectura()) return;
    const monto = Number(data.monto) || 0;
    avisoFondos(monto, `anticipo a ${empleadoName(data.empleadoId)}`);
    const { data: nuevo, error } = await supabase.from('anticipos').insert({
      empresa_id: empresaActiva, empleado_id: data.empleadoId, fecha: data.fecha, monto, motivo: data.motivo, estado: 'pendiente',
    }).select().single();
    if (error) { alert('No se pudo guardar el anticipo: ' + error.message); return; }
    setAnticipos((p) => [...p, { id: nuevo.id, ...data, monto, estado: 'pendiente' }]);
    postAsiento(data.fecha, `Anticipo de sueldo — ${empleadoName(data.empleadoId)}`, '1310', monto, '1110', 'CC-ADM');
  };
  // ---------- Préstamos a empleados ----------
  const addPrestamo = async (data) => {
    if (bloqueadoPorLectura()) return;
    const montoTotal = Number(data.montoTotal) || 0;
    const cuotas = Number(data.cuotas) || 1;
    avisoFondos(montoTotal, `préstamo a ${empleadoName(data.empleadoId)}`);
    const montoCuota = montoTotal / cuotas;
    const { data: nuevo, error } = await supabase.from('prestamos').insert({
      empresa_id: empresaActiva, empleado_id: data.empleadoId, fecha: data.fecha, motivo: data.motivo,
      monto_total: montoTotal, cuotas, monto_cuota: montoCuota, saldo_pendiente: montoTotal, cuotas_pagadas: 0, estado: 'activo',
    }).select().single();
    if (error) { alert('No se pudo guardar el préstamo: ' + error.message); return; }
    setPrestamos((p) => [...p, { id: nuevo.id, empleadoId: data.empleadoId, fecha: data.fecha, motivo: data.motivo, montoTotal, cuotas, montoCuota, saldoPendiente: montoTotal, cuotasPagadas: 0, estado: 'activo' }]);
    postAsiento(data.fecha, `Préstamo a empleado — ${empleadoName(data.empleadoId)}`, '1315', montoTotal, '1110', 'CC-ADM');
  };

  const addNomina = async (data) => {
    if (bloqueadoPorLectura()) return;
    const emp = empleados.find((e) => e.id === data.empleadoId);
    const bruto = Number(data.monto) || 0;
    // Descuento automático: anticipos pendientes de este empleado (completo) y una cuota de préstamo activo, si existen
    const anticiposPend = anticipos.filter((a) => a.empleadoId === data.empleadoId && a.estado === 'pendiente');
    const descuentoAnticipo = anticiposPend.reduce((s, a) => s + a.monto, 0);
    const prestamoActivo = prestamos.find((p) => p.empleadoId === data.empleadoId && p.estado === 'activo');
    const descuentoPrestamo = prestamoActivo ? Math.min(prestamoActivo.montoCuota, prestamoActivo.saldoPendiente) : 0;
    const neto = Math.max(0, bruto - descuentoAnticipo - descuentoPrestamo);
    const { data: nueva, error } = await supabase.from('nomina').insert({
      empresa_id: empresaActiva, empleado_id: data.empleadoId, periodo: data.periodo, fecha_pago: data.fechaPago,
      monto_bruto: bruto, descuento_anticipo: descuentoAnticipo, descuento_prestamo: descuentoPrestamo, monto: neto, estado: 'pendiente',
    }).select().single();
    if (error) { alert('No se pudo guardar la nómina: ' + error.message); return; }
    setNomina((p) => [...p, { id: nueva.id, estado: 'pendiente', ...data, montoBruto: bruto, descuentoAnticipo, descuentoPrestamo, monto: neto }]);
    // Técnicos → Costo de mano de obra (6140); administrativos → Gasto de nómina (5105)
    const cuentaGasto = emp?.rol === 'Técnico' ? '6140' : '5105';
    postAsiento(data.fechaPago, `Nómina ${data.periodo} — ${empleadoName(data.empleadoId)}`, cuentaGasto, bruto, '2510', 'CC-ADM');
    // Liquidación de anticipos y préstamo contra la nómina por pagar (ya no se pagan dos veces)
    if (descuentoAnticipo > 0) {
      postAsiento(data.fechaPago, `Descuento de anticipo en nómina — ${empleadoName(data.empleadoId)}`, '2510', descuentoAnticipo, '1310', 'CC-ADM');
      setAnticipos((prev) => prev.map((a) => a.empleadoId === data.empleadoId && a.estado === 'pendiente' ? { ...a, estado: 'descontado' } : a));
      const idsAnticipos = anticiposPend.map((a) => a.id);
      if (idsAnticipos.length) supabase.from('anticipos').update({ estado: 'descontado' }).in('id', idsAnticipos).then(({ error }) => { if (error) console.error(error); });
    }
    if (descuentoPrestamo > 0 && prestamoActivo) {
      postAsiento(data.fechaPago, `Descuento de cuota de préstamo en nómina — ${empleadoName(data.empleadoId)}`, '2510', descuentoPrestamo, '1315', 'CC-ADM');
      const nuevoSaldo = prestamoActivo.saldoPendiente - descuentoPrestamo;
      const nuevoEstadoPrestamo = nuevoSaldo <= 0.01 ? 'liquidado' : 'activo';
      setPrestamos((prev) => prev.map((p) => p.id === prestamoActivo.id ? { ...p, saldoPendiente: nuevoSaldo, cuotasPagadas: p.cuotasPagadas + 1, estado: nuevoEstadoPrestamo } : p));
      supabase.from('prestamos').update({ saldo_pendiente: nuevoSaldo, cuotas_pagadas: prestamoActivo.cuotasPagadas + 1, estado: nuevoEstadoPrestamo }).eq('id', prestamoActivo.id)
        .then(({ error }) => { if (error) console.error(error); });
    }
  };
  const registrarNominaMes = (mes, ciclo = 'mensual') => {
    const conSalario = empleados.filter((e) => Number(e.salarioMensual) > 0 && e.activo !== false);
    const periodo = ciclo === 'mensual' ? mes : `${mes}-${ciclo === 'quincena1' ? 'Q1' : 'Q2'}`;
    const pendientes = conSalario.filter((e) => !nomina.some((n) => n.empleadoId === e.id && n.periodo === periodo));
    if (pendientes.length === 0) { alert(`La nómina de ${periodo} ya está registrada para todos los empleados con salario.`); return; }
    const montoPorEmpleado = (e) => ciclo === 'mensual' ? Number(e.salarioMensual) : Number(e.salarioMensual) / 2;
    const [anio, mesNum] = mes.split('-').map(Number);
    const ultimoDia = new Date(anio, mesNum, 0).getDate();
    const fechaPago = ciclo === 'quincena1' ? `${mes}-15` : `${mes}-${String(ultimoDia).padStart(2, '0')}`;
    pedirConfirmacion(
      `Se registrará la nómina de ${periodo} para ${pendientes.length} empleado(s):\n${pendientes.map((e) => `• ${e.nombre}: ${fmt(montoPorEmpleado(e))}`).join('\n')}\n\nTotal: ${fmt(pendientes.reduce((s, e) => s + montoPorEmpleado(e), 0))}`,
      () => {
        pendientes.forEach((e) => addNomina({ empleadoId: e.id, periodo, fechaPago, monto: montoPorEmpleado(e) }));
        alert(`Nómina de ${periodo} registrada: ${pendientes.length} empleado(s).`);
      }
    );
  };
  const toggleNominaPago = (id) => {
    if (bloqueadoPorLectura()) return;
    const nom = nomina.find((n) => n.id === id);
    if (nom && nom.estado !== 'pagado') avisoFondos(nom.monto, `pago nómina ${nom.periodo}`);
    const nuevo = nom?.estado === 'pagado' ? 'pendiente' : 'pagado';
    setNomina((prev) => prev.map((n) => {
      if (n.id !== id) return n;
      const hoy = new Date().toISOString().slice(0, 10);
      if (nuevo === 'pagado') {
        postAsiento(hoy, `Pago de nómina — ${empleadoName(n.empleadoId)} (${n.periodo})`, '2510', n.monto, '1110', 'CC-ADM');
      } else {
        // Reversión: asiento inverso para no dejar el pago colgado en la contabilidad
        postAsiento(hoy, `Reversión pago de nómina — ${empleadoName(n.empleadoId)} (${n.periodo})`, '1110', n.monto, '2510', 'CC-ADM');
      }
      return { ...n, estado: nuevo };
    }));
    supabase.from('nomina').update({ estado: nuevo }).eq('id', id).then(({ error }) => { if (error) console.error(error); });
  };
  const nominaMensual = empleados.reduce((s, e) => s + e.salarioMensual, 0);
  const nominaPendiente = nomina.filter((n) => n.estado === 'pendiente').reduce((s, n) => s + n.monto, 0);

  // ---------- Contabilidad general ----------
  const cuentaNombre = (codigo) => cuentas.find((c) => c.codigo === codigo)?.nombre || codigo;
  const addAsiento = (data) => {
    if (bloqueadoPorLectura()) return;
    if (cierreFecha && data.fecha <= cierreFecha) { alert(`El periodo hasta ${cierreFecha} está cerrado. No se pueden registrar comprobantes con esa fecha.`); return; }
    const nuevo = { id: nextId('AS', asientosRef.current), ...data, montoDebe: Number(data.montoDebe) || 0, montoHaber: Number(data.montoHaber) || 0 };
    asientosRef.current = [...asientosRef.current, nuevo];
    setAsientos(asientosRef.current);
    if (empresaActiva) {
      const cD = cuentas.find((c) => c.codigo === data.cuentaDebe)?.id;
      const cH = cuentas.find((c) => c.codigo === data.cuentaHaber)?.id;
      const ctro = data.centroCostoId ? centros.find((c) => c.id === data.centroCostoId)?.id : null;
      if (cD && cH) {
        supabase.from('asientos').insert({
          empresa_id: empresaActiva, fecha: data.fecha, glosa: data.glosa,
          cuenta_debe_id: cD, monto_debe: nuevo.montoDebe, cuenta_haber_id: cH, monto_haber: nuevo.montoHaber,
          centro_costo_id: ctro || null, usuario_id: sesion?.id || null,
        }).then(({ error }) => { if (error) console.error('Error al guardar comprobante:', error.message); });
      }
    }
  };
  const updateAsiento = (id, data) => {
    const original = asientos.find((a) => a.id === id);
    if (cierreFecha && original && (original.fecha <= cierreFecha || data.fecha <= cierreFecha)) { alert(`El periodo hasta ${cierreFecha} está cerrado. Este comprobante no se puede modificar.`); return; }
    asientosRef.current = asientosRef.current.map((a) => a.id === id ? { ...a, ...data, montoDebe: Number(data.montoDebe) || 0, montoHaber: Number(data.montoHaber) || 0 } : a);
    setAsientos(asientosRef.current);
    if (empresaActiva && original?.supabaseId) {
      const cD = cuentas.find((c) => c.codigo === data.cuentaDebe)?.id;
      const cH = cuentas.find((c) => c.codigo === data.cuentaHaber)?.id;
      supabase.from('asientos').update({
        fecha: data.fecha, glosa: data.glosa, cuenta_debe_id: cD, monto_debe: Number(data.montoDebe) || 0,
        cuenta_haber_id: cH, monto_haber: Number(data.montoHaber) || 0,
      }).eq('id', original.supabaseId).then(({ error }) => { if (error) console.error('Error al editar comprobante:', error.message); });
    }
  };
  const anularAsiento = (id) => {
    const original = asientos.find((a) => a.id === id);
    if (cierreFecha && original && original.fecha <= cierreFecha) { alert(`El periodo hasta ${cierreFecha} está cerrado. Este comprobante no se puede anular.`); return; }
    pedirConfirmacion('¿Anular este comprobante? Dejará de afectar todos los reportes contables, pero quedará visible en el libro diario como anulado.', () => {
      asientosRef.current = asientosRef.current.map((a) => a.id === id ? { ...a, anulado: true } : a);
      setAsientos(asientosRef.current);
      if (original?.supabaseId) {
        supabase.from('asientos').update({ anulado: true }).eq('id', original.supabaseId)
          .then(({ error }) => { if (error) console.error('Error al anular comprobante:', error.message); });
      }
    });
  };
  // Asientos vigentes: excluye los anulados de TODOS los cálculos contables
  const asientosVigentes = asientos.filter((a) => !a.anulado);
  // Identifica el centro Administrativo sin importar si el valor guardado es el código 'CC-ADM' (asientos) o el UUID real (activos/órdenes, cargados desde Supabase).
  const esCentroAdmin = (valor) => valor === 'CC-ADM' || centros.find((c) => c.id === valor)?.codigo === 'CC-ADM';

  // ---------- Conciliación bancaria: importación de estado de cuenta y matching automático ----------
  const movsBancoLibro = () => asientosVigentes
    .filter((a) => (a.cuentaDebe === '1110' || a.cuentaHaber === '1110') && a.fecha <= conciliacionCorte)
    .map((a) => ({ ...a, monto: a.cuentaDebe === '1110' ? a.montoDebe : -a.montoHaber }));

  const parseEstadoCuenta = (text) => {
    const rows = [];
    text.split(/\r?\n/).forEach((line) => {
      if (!line.trim()) return;
      const parts = line.split(/[;,\t]/).map((p) => p.trim().replace(/^"|"$/g, ''));
      if (parts.length < 2) return;
      const monto = Number(String(parts[parts.length - 1]).replace(/[$\s]/g, '').replace(/,/g, ''));
      if (isNaN(monto)) return; // encabezado u otra línea no numérica
      let fecha = parts[0];
      const m = fecha.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
      if (m) fecha = `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
      const descripcion = parts.slice(1, parts.length - 1).join(' ') || 'Movimiento bancario';
      rows.push({ id: 'B' + String(rows.length + 1).padStart(3, '0'), fecha, descripcion, monto, estado: 'pendiente' });
    });
    return rows;
  };

  const importarEstadoCuenta = (file) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const rows = parseEstadoCuenta(String(e.target.result));
      if (rows.length === 0) { alert('No se encontraron movimientos válidos. Formato esperado por línea: fecha, descripción, monto (monto negativo = cargo).'); return; }
      setBancoMovs(rows);
    };
    reader.readAsText(file);
  };

  const diasEntre = (f1, f2) => Math.abs((new Date(f1) - new Date(f2)) / (1000 * 60 * 60 * 24));
  const conciliarAutomatico = () => {
    const libro = movsBancoLibro();
    const nuevos = { ...conciliados };
    const usadosAhora = new Set();
    const resultado = bancoMovs.map((b) => {
      const match = libro.find((l) => !nuevos[l.id] && !usadosAhora.has(l.id) && Math.abs(l.monto - b.monto) < 0.01 && diasEntre(l.fecha, b.fecha) <= 7);
      if (match) { nuevos[match.id] = true; usadosAhora.add(match.id); return { ...b, estado: 'conciliado', asientoId: match.id }; }
      return { ...b, estado: 'sin_coincidencia' };
    });
    setConciliados(nuevos);
    setBancoMovs(resultado);
  };

  const guardarConciliacion = (snapshot) => {
    setConciliacionesGuardadas((prev) => [...prev, { id: 'CN' + String(prev.length + 1).padStart(3, '0'), ...snapshot, guardadoEl: new Date().toISOString().slice(0, 16).replace('T', ' ') }]);
  };
  const updateCuentaNombre = (codigo, nombre) => {
    setCuentas((prev) => prev.map((c) => c.codigo === codigo ? { ...c, nombre } : c));
    const id = cuentas.find((c) => c.codigo === codigo)?.id;
    if (id) supabase.from('cuentas').update({ nombre }).eq('id', id).then(({ error }) => { if (error) console.error(error); });
  };
  const updateCuentaSubtipo = (codigo, subtipo) => {
    setCuentas((prev) => prev.map((c) => c.codigo === codigo ? { ...c, subtipo: subtipo || null } : c));
    const id = cuentas.find((c) => c.codigo === codigo)?.id;
    if (id) supabase.from('cuentas').update({ subtipo: subtipo || null }).eq('id', id).then(({ error }) => { if (error) console.error(error); });
  };
  const addCuenta = async (data) => {
    if (cuentas.some((c) => c.codigo === data.codigo)) return;
    const { data: nueva, error } = await supabase.from('cuentas').insert({
      empresa_id: empresaActiva, codigo: data.codigo, nombre: data.nombre, tipo: data.tipo, subtipo: data.subtipo || null, grupo: data.grupo || null,
    }).select().single();
    if (error) { alert('No se pudo guardar la cuenta: ' + error.message); return; }
    setCuentas((p) => [...p, { id: nueva.id, codigo: nueva.codigo, nombre: nueva.nombre, tipo: nueva.tipo, subtipo: nueva.subtipo, grupo: nueva.grupo }].sort((a, b) => a.codigo.localeCompare(b.codigo)));
  };
  const updateCuentaGrupo = (codigo, grupo) => {
    setCuentas((prev) => prev.map((c) => c.codigo === codigo ? { ...c, grupo: grupo || null } : c));
    const id = cuentas.find((c) => c.codigo === codigo)?.id;
    if (id) supabase.from('cuentas').update({ grupo: grupo || null }).eq('id', id).then(({ error }) => { if (error) console.error(error); });
  };
  const deleteCuenta = (codigo) => {
    const enUso = asientos.some((a) => a.cuentaDebe === codigo || a.cuentaHaber === codigo);
    if (enUso) { alert('Esta cuenta ya tiene movimientos registrados y no se puede eliminar.'); return; }
    const id = cuentas.find((c) => c.codigo === codigo)?.id;
    setCuentas((prev) => prev.filter((c) => c.codigo !== codigo));
    if (id) supabase.from('cuentas').delete().eq('id', id).then(({ error }) => { if (error) console.error(error); });
  };

  // ---------- Parametrización: centros de costo y categorías ----------
  const addCentro = async (data) => {
    if (centros.some((c) => c.codigo === data.id)) { alert('Ya existe un centro con ese código.'); return; }
    const { data: nuevo, error } = await supabase.from('centros_costo').insert({ empresa_id: empresaActiva, codigo: data.id, nombre: data.nombre }).select().single();
    if (error) { alert('No se pudo guardar el centro de costo: ' + error.message); return; }
    setCentros((p) => [...p, { id: nuevo.id, codigo: nuevo.codigo, nombre: nuevo.nombre }]);
  };
  const deleteCentro = (id) => {
    const enUso = asientos.some((a) => a.centroCostoId === id) || orders.some((o) => o.centroCostoId === id) || activos.some((a) => a.centroCostoId === id);
    if (enUso) { alert('Este centro de costo ya tiene movimientos u objetos asignados y no se puede eliminar.'); return; }
    setCentros((prev) => prev.filter((c) => c.id !== id));
    supabase.from('centros_costo').delete().eq('id', id).then(({ error }) => { if (error) console.error(error); });
  };
  const addCategoria = async (tipo, data) => {
    const lista = tipo === 'gasto' ? gastoCategorias : ingresoCategorias;
    if (lista.some((c) => c.nombre === data.nombre)) { alert('Ya existe una categoría con ese nombre.'); return; }
    const cuentaId = cuentas.find((c) => c.codigo === data.cuenta)?.id;
    if (!cuentaId) { alert('No se encontró la cuenta seleccionada.'); return; }
    const { error } = await supabase.from('categorias').insert({ empresa_id: empresaActiva, tipo, nombre: data.nombre, cuenta_id: cuentaId });
    if (error) { alert('No se pudo guardar la categoría: ' + error.message); return; }
    const setter = tipo === 'gasto' ? setGastoCategorias : setIngresoCategorias;
    setter((p) => [...p, { nombre: data.nombre, cuenta: data.cuenta }]);
  };
  const updateCategoriaCuenta = (tipo, nombre, cuenta) => {
    const setter = tipo === 'gasto' ? setGastoCategorias : setIngresoCategorias;
    setter((prev) => prev.map((c) => c.nombre === nombre ? { ...c, cuenta } : c));
    const cuentaId = cuentas.find((c) => c.codigo === cuenta)?.id;
    if (cuentaId) supabase.from('categorias').update({ cuenta_id: cuentaId }).eq('empresa_id', empresaActiva).eq('tipo', tipo).eq('nombre', nombre)
      .then(({ error }) => { if (error) console.error(error); });
  };
  const deleteCategoria = (tipo, nombre) => {
    const enUso = movimientos.some((m) => m.categoria === nombre)
      || (tipo === 'gasto' && cxp.some((c) => c.categoria === nombre))
      || (tipo === 'ingreso' && facturasManuales.some((f) => f.categoria === nombre));
    if (enUso) { alert('Esta categoría ya tiene movimientos registrados y no se puede eliminar.'); return; }
    const setter = tipo === 'gasto' ? setGastoCategorias : setIngresoCategorias;
    setter((prev) => prev.filter((c) => c.nombre !== nombre));
    supabase.from('categorias').delete().eq('empresa_id', empresaActiva).eq('tipo', tipo).eq('nombre', nombre).then(({ error }) => { if (error) console.error(error); });
  };

  const mayorRows = asientosVigentes
    .filter((a) => a.cuentaDebe === cuentaMayor || a.cuentaHaber === cuentaMayor)
    .filter((a) => mayorCentro === 'todos' || a.centroCostoId === mayorCentro)
    .sort((a, b) => a.fecha.localeCompare(b.fecha))
    .reduce((acc, a) => {
      const debe = a.cuentaDebe === cuentaMayor ? a.montoDebe : 0;
      const haber = a.cuentaHaber === cuentaMayor ? a.montoHaber : 0;
      const prevSaldo = acc.length ? acc[acc.length - 1].saldo : 0;
      acc.push({ ...a, debe, haber, saldo: prevSaldo + debe - haber });
      return acc;
    }, []);
  const balanceComprobacion = cuentas.map((c) => {
    const debe = asientosVigentes.filter((a) => a.cuentaDebe === c.codigo).reduce((s, a) => s + a.montoDebe, 0);
    const haber = asientosVigentes.filter((a) => a.cuentaHaber === c.codigo).reduce((s, a) => s + a.montoHaber, 0);
    return { ...c, debe, haber };
  }).filter((c) => c.debe > 0 || c.haber > 0);
  const totalDebe = balanceComprobacion.reduce((s, c) => s + c.debe, 0);
  const totalHaber = balanceComprobacion.reduce((s, c) => s + c.haber, 0);

  // ---------- Balance general (a la fecha "Al" del filtro) ----------
  const asOfDate = reportTo || '9999-12-31';
  const asientosAsOf = asientosVigentes.filter((a) => a.fecha <= asOfDate);
  const cuentaSaldoAsOf = (codigo, tipo) => {
    const debe = asientosAsOf.filter((a) => a.cuentaDebe === codigo).reduce((s, a) => s + a.montoDebe, 0);
    const haber = asientosAsOf.filter((a) => a.cuentaHaber === codigo).reduce((s, a) => s + a.montoHaber, 0);
    return (tipo === 'Activo' || tipo === 'Costo' || tipo === 'Gasto') ? debe - haber : haber - debe;
  };
  // Ingresos/Costos/Utilidad del Panel general: se leen del libro mayor (igual que el Balance General), no se
  // recalculan aparte desde las órdenes — así los tres reportes (Panel, Estado de Resultados, Balance) parten
  // siempre de la misma fuente y nunca quedan descuadrados entre sí. La única diferencia real y esperada es que
  // el Estado de Resultados sí filtra por el rango "Del/Al" que elijas ahí (es "de un periodo"), mientras que
  // el Panel general y el Balance son acumulados a la fecha de corte (asOfDate) — como debe ser un balance.
  const revenue = useMemo(() => cuentas.filter((c) => c.tipo === 'Ingreso').reduce((s, c) => s + cuentaSaldoAsOf(c.codigo, c.tipo), 0), [asientos, cuentas, asOfDate]);
  const costs = useMemo(() => cuentas.filter((c) => c.tipo === 'Costo' || c.tipo === 'Gasto').reduce((s, c) => s + cuentaSaldoAsOf(c.codigo, c.tipo), 0), [asientos, cuentas, asOfDate]);
  const profit = revenue - costs;
  const cuentasActivo = cuentas.filter((c) => c.tipo === 'Activo').map((c) => ({ ...c, saldo: cuentaSaldoAsOf(c.codigo, c.tipo) }));
  const cuentasPasivo = cuentas.filter((c) => c.tipo === 'Pasivo').map((c) => ({ ...c, saldo: cuentaSaldoAsOf(c.codigo, c.tipo) }));
  const cuentasCapital = cuentas.filter((c) => c.tipo === 'Capital').map((c) => ({ ...c, saldo: cuentaSaldoAsOf(c.codigo, c.tipo) }));
  const utilidadEjercicio = cuentas.filter((c) => c.tipo === 'Ingreso').reduce((s, c) => s + cuentaSaldoAsOf(c.codigo, c.tipo), 0)
    - cuentas.filter((c) => c.tipo === 'Costo' || c.tipo === 'Gasto').reduce((s, c) => s + cuentaSaldoAsOf(c.codigo, c.tipo), 0);
  const totalActivo = cuentasActivo.reduce((s, c) => s + c.saldo, 0);
  const totalPasivo = cuentasPasivo.reduce((s, c) => s + c.saldo, 0);
  const totalCapital = cuentasCapital.reduce((s, c) => s + c.saldo, 0) + utilidadEjercicio;
  const totalPasivoCapital = totalPasivo + totalCapital;

  // ---------- Flujo de efectivo (del periodo Del/Al, cuentas Caja 1105 y Bancos 1110) ----------
  const CASH_ACCOUNTS = ['1105', '1110'];
  const esCash = (codigo) => CASH_ACCOUNTS.includes(codigo);
  const clasificaFlujo = (codigo) => codigo === '1540' ? 'inversion' : codigo === '3105' ? 'financiamiento' : 'operacion';
  // Excluir traspasos internos Caja<->Bancos (no son flujo, solo mueven efectivo de bolsillo)
  const cashAsientos = asientosVigentes.filter((a) => (esCash(a.cuentaDebe) || esCash(a.cuentaHaber)) && !(esCash(a.cuentaDebe) && esCash(a.cuentaHaber)));
  const saldoInicialCaja = cashAsientos.filter((a) => a.fecha < reportFrom).reduce((s, a) => s + (esCash(a.cuentaDebe) ? a.montoDebe : -a.montoHaber), 0);
  const flujoDetalle = { operacion: [], inversion: [], financiamiento: [] };
  cashAsientos.filter((a) => inRange(a.fecha)).forEach((a) => {
    const entra = esCash(a.cuentaDebe);
    const monto = entra ? a.montoDebe : -a.montoHaber;
    const cat = clasificaFlujo(entra ? a.cuentaHaber : a.cuentaDebe);
    flujoDetalle[cat].push({ ...a, monto });
  });
  const flujoOperacion = flujoDetalle.operacion.reduce((s, a) => s + a.monto, 0);
  const flujoInversion = flujoDetalle.inversion.reduce((s, a) => s + a.monto, 0);
  const flujoFinanciamiento = flujoDetalle.financiamiento.reduce((s, a) => s + a.monto, 0);
  const flujoNetoPeriodo = flujoOperacion + flujoInversion + flujoFinanciamiento;
  const saldoFinalCaja = saldoInicialCaja + flujoNetoPeriodo;

  // ---------- Estado de resultados basado en el libro mayor (agrupado por cuenta) ----------
  const erAsientos = asientosVigentes.filter((a) => inRange(a.fecha) && (estadoResCentro === 'todos' || a.centroCostoId === estadoResCentro));
  const cuentaSaldoPeriodo = (codigo, tipo) => {
    const debe = erAsientos.filter((a) => a.cuentaDebe === codigo).reduce((s, a) => s + a.montoDebe, 0);
    const haber = erAsientos.filter((a) => a.cuentaHaber === codigo).reduce((s, a) => s + a.montoHaber, 0);
    return tipo === 'Ingreso' ? haber - debe : debe - haber;
  };
  const erGrupos = ER_GRUPOS.map((g) => {
    const cuentasGrupo = cuentas.filter((c) => c.grupo === g.key).map((c) => ({ ...c, saldo: cuentaSaldoPeriodo(c.codigo, c.tipo) }));
    return { ...g, cuentas: cuentasGrupo, total: cuentasGrupo.reduce((s, c) => s + c.saldo, 0) };
  });
  const erGrupo = (key) => erGrupos.find((g) => g.key === key) || { cuentas: [], total: 0 };
  const erIngresos = erGrupo('ingresos_operativos').total;
  const erCostoMat = erGrupo('costo_materiales').total;
  const erCostoMO = erGrupo('costo_mano_obra').total;
  const erCostoDep = erGrupo('costo_depreciacion').total;
  const erOtrosIngresos = erGrupo('otros_ingresos').total;
  const erGastos = erGrupo('gastos_generales').total;
  // Prorrateo de gastos administrativos integrado EN la línea de Gastos generales cuando se ve un centro operativo
  const esCentroOperativo = estadoResCentro !== 'todos' && estadoResCentro !== 'CC-ADM';
  const erProrrateoInfo = (() => {
    if (!esCentroOperativo) return { monto: 0, factor: 0 };
    const admAsientos = asientosVigentes.filter((a) => inRange(a.fecha) && esCentroAdmin(a.centroCostoId));
    const saldoAdm = (codigo, tipo) => {
      const d = admAsientos.filter((a) => a.cuentaDebe === codigo).reduce((s, a) => s + a.montoDebe, 0);
      const h = admAsientos.filter((a) => a.cuentaHaber === codigo).reduce((s, a) => s + a.montoHaber, 0);
      return tipo === 'Ingreso' ? h - d : d - h;
    };
    const gastosAdmTotal = cuentas.filter((c) => c.grupo === 'gastos_generales').reduce((s, c) => s + saldoAdm(c.codigo, c.tipo), 0);
    const centrosOp = centros.filter((c) => c.codigo !== 'CC-ADM');
    let factor;
    if (prorrateoMetodo === 'equitativo') {
      factor = centrosOp.length > 0 ? 1 / centrosOp.length : 0;
    } else {
      const ingTotales = asientosVigentes.filter((a) => inRange(a.fecha) && !esCentroAdmin(a.centroCostoId) && cuentas.find((c) => c.codigo === a.cuentaHaber)?.grupo === 'ingresos_operativos').reduce((s, a) => s + a.montoHaber, 0)
        - asientosVigentes.filter((a) => inRange(a.fecha) && !esCentroAdmin(a.centroCostoId) && cuentas.find((c) => c.codigo === a.cuentaDebe)?.grupo === 'ingresos_operativos').reduce((s, a) => s + a.montoDebe, 0);
      factor = ingTotales > 0 ? erIngresos / ingTotales : (centrosOp.length > 0 ? 1 / centrosOp.length : 0);
    }
    return { monto: gastosAdmTotal * factor, factor };
  })();
  const erGastosConProrrateo = erGastos + erProrrateoInfo.monto;
  const erUtilidadOperativa = erIngresos - erCostoMat - erCostoMO - erCostoDep;
  const erUtilidadNeta = erUtilidadOperativa + erOtrosIngresos - erGastosConProrrateo;
  // Cuentas de Ingreso/Costo/Gasto sin grupo asignado (para avisar que falta configurarlas)
  const cuentasSinGrupo = cuentas.filter((c) => ['Ingreso', 'Costo', 'Gasto'].includes(c.tipo) && !c.grupo);
  const erLineas = [
    { label: 'Ingresos por instalaciones', total: erIngresos, cuentas: erGrupo('ingresos_operativos').cuentas },
    { label: '(-) Costo de materiales', total: -erCostoMat, cuentas: erGrupo('costo_materiales').cuentas, resta: true },
    { label: '(-) Costo de mano de obra', total: -erCostoMO, cuentas: erGrupo('costo_mano_obra').cuentas, resta: true },
    { label: '(-) Costo de depreciación', total: -erCostoDep, cuentas: erGrupo('costo_depreciacion').cuentas, resta: true },
  ];
  const erLineasCorporativas = [
    { label: '(+) Otros ingresos', total: erOtrosIngresos, cuentas: erGrupo('otros_ingresos').cuentas },
    {
      label: '(-) Gastos generales', total: -erGastosConProrrateo, resta: true,
      cuentas: esCentroOperativo && erProrrateoInfo.monto > 0
        ? [...erGrupo('gastos_generales').cuentas, { codigo: '····', nombre: `Gastos administrativos prorrateados (${prorrateoMetodo === 'ventas' ? 'por ventas' : 'equitativo'}, ${(erProrrateoInfo.factor * 100).toFixed(1)}%)`, saldo: erProrrateoInfo.monto }]
        : erGrupo('gastos_generales').cuentas,
    },
  ];

  // ---------- ER por rango arbitrario (para acumulado y comparativo) ----------
  const erTotalsFor = (from, to, centro) => {
    const enRango = (f) => (!from || f >= from) && (!to || f <= to);
    const filtrados = asientosVigentes.filter((a) => enRango(a.fecha) && (centro === 'todos' || a.centroCostoId === centro));
    const saldo = (codigo, tipo) => {
      const debe = filtrados.filter((a) => a.cuentaDebe === codigo).reduce((s, a) => s + a.montoDebe, 0);
      const haber = filtrados.filter((a) => a.cuentaHaber === codigo).reduce((s, a) => s + a.montoHaber, 0);
      return tipo === 'Ingreso' ? haber - debe : debe - haber;
    };
    const grupoTotal = (key) => cuentas.filter((c) => c.grupo === key).reduce((s, c) => s + saldo(c.codigo, c.tipo), 0);
    const ingresos = grupoTotal('ingresos_operativos');
    const cMat = grupoTotal('costo_materiales');
    const cMO = grupoTotal('costo_mano_obra');
    const cDep = grupoTotal('costo_depreciacion');
    const otros = grupoTotal('otros_ingresos');
    const gastos = grupoTotal('gastos_generales');
    const utilOp = ingresos - cMat - cMO - cDep;
    // Prorrateo de gastos administrativos (CC-ADM) a los centros operativos
    // Factor por 'ventas' = ingresos por órdenes del centro / ingresos totales; 'equitativo' = 1/n centros operativos
    let gastosAdmProrrateados = 0, factorProrrateo = 0;
    if (centro !== 'todos' && centro !== 'CC-ADM') {
      const filtradosAdm = asientosVigentes.filter((a) => enRango(a.fecha) && esCentroAdmin(a.centroCostoId));
      const saldoAdm = (codigo, tipo) => {
        const d = filtradosAdm.filter((a) => a.cuentaDebe === codigo).reduce((s, a) => s + a.montoDebe, 0);
        const h = filtradosAdm.filter((a) => a.cuentaHaber === codigo).reduce((s, a) => s + a.montoHaber, 0);
        return tipo === 'Ingreso' ? h - d : d - h;
      };
      const gastosAdmTotal = cuentas.filter((c) => c.grupo === 'gastos_generales').reduce((s, c) => s + saldoAdm(c.codigo, c.tipo), 0);
      const centrosOperativos = centros.filter((c) => c.codigo !== 'CC-ADM');
      if (prorrateoMetodo === 'equitativo') {
        factorProrrateo = centrosOperativos.length > 0 ? 1 / centrosOperativos.length : 0;
      } else {
        // Por ventas: ingresos de órdenes del centro / ingresos totales de todos los centros operativos
        const ingresosTotales = asientosVigentes.filter((a) => enRango(a.fecha) && !esCentroAdmin(a.centroCostoId))
          .filter((a) => cuentas.find((c) => c.codigo === a.cuentaHaber)?.grupo === 'ingresos_operativos').reduce((s, a) => s + a.montoHaber, 0)
          - asientosVigentes.filter((a) => enRango(a.fecha) && !esCentroAdmin(a.centroCostoId))
          .filter((a) => cuentas.find((c) => c.codigo === a.cuentaDebe)?.grupo === 'ingresos_operativos').reduce((s, a) => s + a.montoDebe, 0);
        factorProrrateo = ingresosTotales > 0 ? ingresos / ingresosTotales : (centrosOperativos.length > 0 ? 1 / centrosOperativos.length : 0);
      }
      gastosAdmProrrateados = gastosAdmTotal * factorProrrateo;
    }
    const gastosConProrrateo = gastos + gastosAdmProrrateados;
    return { ingresos, cMat, cMO, cDep, utilOp, otros, gastos: gastosConProrrateo, utilNeta: utilOp + otros - gastosConProrrateo, gastosAdmProrrateados, factorProrrateo, utilNetaCentro: utilOp + otros - gastosConProrrateo };
  };

  // ---------- Exportación CSV de informes ----------
  const descargarCSVInforme = (nombre, rows) => {
    const csv = rows.map((r) => r.map((cell) => `"${String(cell ?? '').replace(/"/g, '""')}"`).join(',')).join('\n');
    const blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = `${nombre}.csv`; a.click();
    URL.revokeObjectURL(url);
  };
  // Respaldo completo del sistema: TODAS las empresas y datos globales en un solo JSON.
  // Sirve como copia de seguridad y como archivo de origen para poblar Supabase el día de la migración.
  // NOTA: estas 3 funciones vivían de la simulación local de multiempresa (ya retirada).
  // Con datos reales en Supabase, "respaldo/restaurar/iniciar de cero" necesitan un diseño nuevo
  // (exportar/importar filas reales de la base, y un borrado real y controlado). Pendiente de próxima ronda.
  const descargarRespaldoCompleto = () => {
    alert('El respaldo completo se está rediseñando para trabajar con la base de datos real de Supabase. Por ahora, los respaldos se hacen desde el panel de Supabase (Database → Backups).');
  };
  const restaurarRespaldo = (file) => {
    alert('Restaurar respaldo está pendiente de rediseño para la base de datos real. Por ahora, cualquier restauración se hace directo en Supabase.');
  };
  const iniciarDeCero = () => {
    alert('"Iniciar de cero" está pendiente de rediseño: ahora borraría datos reales de tu base en Supabase, así que merece su propio flujo cuidadoso en vez de un botón genérico. Por ahora, pídeme a mí que lo haga directo en la base si lo necesitas.');
  };
  const exportInformeActual = () => {
    if (contabVista === 'balance_general') {
      const rows = [['Balance General al ' + (reportTo || 'hoy')], ['Sección', 'Cuenta', 'Saldo']];
      cuentasActivo.forEach((c) => rows.push([BG_SUBTIPOS.find((s) => s.key === c.subtipo)?.label || 'Activo sin clasificar', c.nombre, c.saldo]));
      rows.push(['', 'TOTAL ACTIVO', totalActivo]);
      cuentasPasivo.forEach((c) => rows.push([BG_SUBTIPOS.find((s) => s.key === c.subtipo)?.label || 'Pasivo sin clasificar', c.nombre, c.saldo]));
      rows.push(['', 'TOTAL PASIVO', totalPasivo]);
      cuentasCapital.forEach((c) => rows.push(['Capital', c.nombre, c.saldo]));
      rows.push(['Capital', 'Utilidad del ejercicio', utilidadEjercicio], ['', 'TOTAL CAPITAL', totalCapital], ['', 'TOTAL PASIVO + CAPITAL', totalPasivoCapital]);
      descargarCSVInforme('balance-general', rows);
    } else if (contabVista === 'estado_resultados') {
      if (erModo === 'comparativo') {
        const A = erTotalsFor(reportFrom, reportTo, estadoResCentro);
        const B = erTotalsFor(compareFrom, compareTo, estadoResCentro);
        const rows = [['Estado de Resultados comparativo'], ['Concepto', `A: ${reportFrom} a ${reportTo}`, `B: ${compareFrom} a ${compareTo}`, 'Variación $', 'Variación %']];
        [['Ingresos por instalaciones', 'ingresos'], ['Costo de materiales', 'cMat'], ['Costo de mano de obra', 'cMO'], ['Costo de depreciación', 'cDep'], ['Utilidad operativa', 'utilOp'], ['Otros ingresos', 'otros'], ['Gastos generales', 'gastos'], ['Utilidad neta', 'utilNeta']].forEach(([label, k]) => {
          const va = A[k], vb = B[k];
          rows.push([label, va, vb, va - vb, vb !== 0 ? (((va - vb) / Math.abs(vb)) * 100).toFixed(1) + '%' : '—']);
        });
        descargarCSVInforme('estado-resultados-comparativo', rows);
      } else {
        const from = erModo === 'acumulado' ? '' : reportFrom;
        const T = erTotalsFor(from, reportTo, estadoResCentro);
        const rows = [[`Estado de Resultados ${erModo === 'acumulado' ? 'acumulado al ' + reportTo : reportFrom + ' a ' + reportTo}`], ['Concepto', 'Monto'],
          ['Ingresos por instalaciones', T.ingresos], ['(-) Costo de materiales', -T.cMat], ['(-) Costo de mano de obra', -T.cMO], ['(-) Costo de depreciación', -T.cDep],
          ['= Utilidad operativa', T.utilOp], ['(+) Otros ingresos', T.otros], ['(-) Gastos generales', -T.gastos], ['= Utilidad neta', T.utilNeta]];
        descargarCSVInforme('estado-resultados', rows);
      }
    } else if (contabVista === 'balance_comprobacion') {
      const rows = [['Balance de comprobación'], ['Cuenta', 'Debe', 'Haber']];
      balanceComprobacion.forEach((c) => rows.push([`${c.codigo} — ${c.nombre}`, c.debe, c.haber]));
      rows.push(['TOTALES', totalDebe, totalHaber]);
      descargarCSVInforme('balance-comprobacion', rows);
    } else if (contabVista === 'libro_mayor') {
      const rows = [[`Libro mayor — ${cuentaMayor} ${cuentaNombre(cuentaMayor)}`], ['Fecha', 'Glosa', 'Debe', 'Haber', 'Saldo']];
      mayorRows.forEach((r) => rows.push([r.fecha, r.glosa, r.debe || '', r.haber || '', r.saldo]));
      descargarCSVInforme('libro-mayor', rows);
    } else if (contabVista === 'flujo_efectivo') {
      const rows = [['Flujo de efectivo', reportFrom + ' a ' + reportTo], ['Concepto', 'Monto'],
        ['Saldo inicial', saldoInicialCaja], ['Operación', flujoOperacion], ['Inversión', flujoInversion], ['Financiamiento', flujoFinanciamiento], ['Flujo neto', flujoNetoPeriodo], ['Saldo final', saldoFinalCaja]];
      descargarCSVInforme('flujo-efectivo', rows);
    }
  };

  // ---------- KPIs (memoizados: solo recalculan si cambian los datos de origen, no en cada render) ----------
  const activoCorriente = cuentasActivo.filter((c) => c.subtipo === 'activo_corriente').reduce((s, c) => s + c.saldo, 0);
  const pasivoCorriente = cuentasPasivo.filter((c) => c.subtipo === 'pasivo_corriente').reduce((s, c) => s + c.saldo, 0);
  const kpiFin = useMemo(() => ({
    liquidez: pasivoCorriente > 0 ? activoCorriente / pasivoCorriente : null,
    capitalTrabajo: activoCorriente - pasivoCorriente,
    endeudamiento: totalActivo > 0 ? totalPasivo / totalActivo : null,
    margenOperativo: erIngresos > 0 ? erUtilidadOperativa / erIngresos : null,
    margenNeto: erIngresos > 0 ? erUtilidadNeta / erIngresos : null,
    roe: totalCapital > 0 ? erUtilidadNeta / totalCapital : null,
  }), [asientos, cuentas, centros, reportFrom, reportTo]);
  const ordersRango = useMemo(() => orders.filter((o) => inRange(o.fecha)), [orders, reportFrom, reportTo]);
  const kpiOrd = useMemo(() => {
    const ordCompletadas = ordersRango.filter((o) => o.estado === 'completada');
    return {
      total: ordersRango.length,
      completadas: ordCompletadas.length,
      canceladas: ordersRango.filter((o) => o.estado === 'cancelada').length,
      ticketPromedio: ordCompletadas.length > 0 ? ordCompletadas.reduce((s, o) => s + Number(o.precioCliente), 0) / ordCompletadas.length : 0,
      utilidadPromedio: ordCompletadas.length > 0 ? ordCompletadas.reduce((s, o) => s + (Number(o.precioCliente) - orderCost(o)), 0) / ordCompletadas.length : 0,
    };
  }, [ordersRango, materials]);
  const kpiTecnicos = useMemo(() => techs.map((t) => {
    const ots = ordersRango.filter((o) => o.tecnicoId === t.id);
    const comp = ots.filter((o) => o.estado === 'completada');
    return { ...t, ordenes: ots.length, completadas: comp.length, ingresos: comp.reduce((s, o) => s + Number(o.precioCliente), 0), utilidad: comp.reduce((s, o) => s + (Number(o.precioCliente) - orderCost(o)), 0) };
  }).sort((a, b) => b.ingresos - a.ingresos), [ordersRango, empleados, materials]);
  const kpiInv = useMemo(() => {
    const valorInventario = materials.reduce((s, m) => s + m.stock * m.costoUnitario, 0);
    const consumoPeriodo = kardexLog.filter((k) => k.tipo === 'salida' && inRange(k.fecha)).reduce((s, k) => s + k.cantidad * k.costoUnitario, 0);
    return {
      valorInventario,
      consumoPeriodo,
      rotacion: valorInventario > 0 ? consumoPeriodo / valorInventario : null,
      stockBajo: materials.filter((m) => m.stock <= 10),
      topConsumidos: materials.map((m) => ({ ...m, consumido: kardexLog.filter((k) => k.tipo === 'salida' && k.materialId === m.id && inRange(k.fecha)).reduce((s, k) => s + k.cantidad, 0) })).filter((m) => m.consumido > 0).sort((a, b) => b.consumido - a.consumido).slice(0, 5),
    };
  }, [materials, kardexLog, reportFrom, reportTo]);
  const hoyStr = new Date().toISOString().slice(0, 10);
  const diasDesde = (f) => Math.floor((new Date(hoyStr) - new Date(f)) / (1000 * 60 * 60 * 24));
  const cxcPendientes = useMemo(() => orders.filter((o) => o.estado === 'completada' && o.estadoPago === 'pendiente').map((o) => ({ ...o, dias: diasDesde(o.fecha) })), [orders]);
  const cxcTotal = cxcPendientes.reduce((s, o) => s + Number(o.precioCliente), 0);
  const agingCxc = useMemo(() => ({
    '0-30 días': cxcPendientes.filter((o) => o.dias <= 30).reduce((s, o) => s + Number(o.precioCliente), 0),
    '31-60 días': cxcPendientes.filter((o) => o.dias > 30 && o.dias <= 60).reduce((s, o) => s + Number(o.precioCliente), 0),
    'Más de 60 días': cxcPendientes.filter((o) => o.dias > 60).reduce((s, o) => s + Number(o.precioCliente), 0),
  }), [cxcPendientes]);
  const cxpVencidas = useMemo(() => cxp.filter((c) => c.estado === 'pendiente' && c.vencimiento < hoyStr), [cxp]);
  const cxpPorVencer = useMemo(() => cxp.filter((c) => c.estado === 'pendiente' && c.vencimiento >= hoyStr).sort((a, b) => a.vencimiento.localeCompare(b.vencimiento)), [cxp]);

  const nav = [
    {
      key: 'kpi_group', label: t('Indicadores KPI'), icon: LayoutDashboard,
      children: [
        { key: 'dashboard', label: t('Panel general') },
        { key: 'kpi_financieros', label: t('Índices financieros') },
        { key: 'kpi_ordenes', label: t('Gestión de órdenes') },
        { key: 'kpi_tecnicos', label: t('Rendimiento de técnicos') },
        { key: 'kpi_inventario', label: t('Inventario y rotación') },
        { key: 'kpi_cartera', label: t('Cartera y pagos') },
      ],
    },
    {
      key: 'contab_group', label: t('Contabilidad General'), icon: Calculator,
      children: [
        { key: 'contab_plan', label: t('Plan de Cuentas') },
        { key: 'contab_comprobantes', label: t('Comprobantes') },
        { key: 'contab_conciliacion', label: t('Conciliación bancaria') },
        { key: 'contab_informes', label: t('Informes') },
        { key: 'ingresos_gastos', label: t('Ingresos y gastos') },
        { key: 'contab_config', label: t('Configuración') },
      ],
    },
    {
      key: 'ordenes_group', label: t('Órdenes de trabajo'), icon: ClipboardList,
      children: [
        { key: 'ordenes', label: t('Órdenes') },
        { key: 'agenda', label: t('Agenda de técnicos') },
        { key: 'proyectos', label: t('Proyectos') },
      ],
    },
    {
      key: 'activos_group', label: t('Activos fijos'), icon: Truck,
      children: [
        { key: 'activos', label: 'Activos fijos' },
        { key: 'depreciacion', label: 'Control de depreciación' },
        { key: 'flota', label: 'Bitácora de flota' },
      ],
    },
    {
      key: 'cxc_group', label: t('Cuentas por cobrar'), icon: FileText,
      children: [
        { key: 'facturacion', label: t('Facturas') },
        { key: 'clientes', label: t('Clientes') },
      ],
    },
    {
      key: 'cxp_group', label: t('Cuentas por pagar'), icon: Building2,
      children: [
        { key: 'cxp', label: 'Cuentas por pagar' },
        { key: 'proveedores', label: 'Proveedores' },
      ],
    },
    {
      key: 'inv_group', label: t('Inventarios'), icon: Package,
      children: [
        { key: 'inventario', label: t('Materiales') },
        { key: 'kardex', label: t('Kardex') },
      ],
    },
    {
      key: 'rh_group', label: t('Recursos humanos'), icon: UserCheck,
      children: [
        { key: 'rh', label: t('Empleados') },
        { key: 'tecnicos', label: t('Técnicos') },
        { key: 'bonos', label: t('Bonos por productividad') },
      ],
    },
    { key: 'reportes', label: t('Reportes'), icon: BarChart3 },
    { key: 'ventas', label: t('Ventas'), icon: Banknote },
    { key: 'empresas', label: t('Empresas'), icon: Building2 },
    { key: 'usuarios', label: t('Usuarios y roles'), icon: UserCheck },
  ].filter((item) => {
    if (!sesion) return false;
    if (item.key === 'usuarios' || item.key === 'empresas') return sesion.rol === 'Administrador';
    // Módulos por giro: Comercio usa Ventas; los demás giros usan Órdenes/Agenda
    const esComercio = empresaActual?.giro === 'Comercio';
    if (item.key === 'ventas' && !esComercio) return false;
    if (item.key === 'ordenes_group' && esComercio) return false;
    if (item.children) {
      item.children = item.children.filter((c) => tieneAcceso(sesion.rol, c.key));
      return item.children.length > 0;
    }
    return tieneAcceso(sesion.rol, item.key);
  });

  // Si por cualquier motivo la pestaña activa no está permitida para el rol de la sesión, redirige
  // a la primera pantalla que sí tenga permitida — evita que alguien vea contenido de otro módulo por accidente.
  useEffect(() => {
    if (!sesion) return;
    if (rolesConfig[sesion.rol] === undefined) return; // los roles de la empresa aún no cargan; esperar antes de decidir
    const tabsPersonales = ['cuenta', 'preferencias', 'formato_correos'];
    if (tabsPersonales.includes(tab)) return;
    if (tab === 'usuarios' || tab === 'empresas') { if (sesion.rol !== 'Administrador') setTab('agenda'); return; }
    if (!tieneAcceso(sesion.rol, tab)) {
      const primeraDisponible = nav.flatMap((g) => g.children ? g.children.map((c) => c.key) : [g.key])[0];
      setTab(primeraDisponible || 'agenda');
    }
  }, [tab, sesion?.rol, rolesConfig]);

  // Abre automáticamente en el menú el grupo al que pertenece la pantalla activa (en vez de forzar uno fijo)
  useEffect(() => {
    const grupo = nav.find((g) => g.children?.some((c) => c.key === tab));
    if (grupo) setExpandedGroup(grupo.key);
  }, [tab]);

  const ordersVisiblesPorRol = sesion?.rol === 'Tecnico' && sesion?.empleadoId ? orders.filter((o) => o.tecnicoId === sesion.empleadoId) : orders;
  const visibleOrders = (orderFilter === 'todas' ? ordersVisiblesPorRol : ordersVisiblesPorRol.filter((o) => o.estado === orderFilter))
    .filter((o) => {
      if (!ordersSearch.trim()) return true;
      const q = ordersSearch.trim().toLowerCase();
      return o.id.toLowerCase().includes(q)
        || (o.referencia || '').toLowerCase().includes(q)
        || clientName(o.clienteId).toLowerCase().includes(q)
        || techName(o.tecnicoId).toLowerCase().includes(q)
        || (o.direccion || '').toLowerCase().includes(q);
    });
  const ordersPag = usePaginado(visibleOrders, 15);
  const kardexRows = kardexLog.filter((k) => k.materialId === kardexMaterial).sort((a, b) => a.fecha.localeCompare(b.fecha) || a.id.localeCompare(b.id));
  const kardexPag = usePaginado(kardexRows, 20);
  const asientosOrdenados = [...asientos].sort((a, b) => b.fecha.localeCompare(a.fecha));
  const comprobantesPag = usePaginado(asientosOrdenados, 20);
  const ventasOrdenadas = [...ventas].reverse();
  const ventasPag = usePaginado(ventasOrdenadas, 15);
  const cxpOrdenadas = [...cxp].sort((a, b) => a.vencimiento.localeCompare(b.vencimiento));
  const cxpPag = usePaginado(cxpOrdenadas, 15);
  const nominaPag = usePaginado(nomina, 15);
  const kardexMat = matById(kardexMaterial);

  // ---------- Autenticación real con Supabase Auth ----------
  // Construye el objeto de sesión que el resto del sistema espera, a partir del perfil + sus empresas/roles reales.
  const cargarSesionDesdeAuth = async (authUser) => {
    const { data: perfil } = await supabase.from('perfiles').select('*').eq('id', authUser.id).single();
    if (!perfil) return null;
    const { data: accesos } = await supabase
      .from('perfil_empresas')
      .select('empresa_id, roles(nombre, es_administrador)')
      .eq('perfil_id', authUser.id);
    const empresasIds = (accesos || []).map((a) => a.empresa_id);
    const esAdministradorGlobal = (accesos || []).some((a) => a.roles?.es_administrador);
    // Rol "de referencia" para compatibilidad con el resto del sistema: el de la empresa activa si ya se conoce, si no el primero
    const rolEnEmpresa = (empId) => accesos?.find((a) => a.empresa_id === empId)?.roles?.nombre || accesos?.[0]?.roles?.nombre || 'Técnico';
    return {
      id: authUser.id,
      usuario: authUser.email,
      username: perfil.username,
      nombre: perfil.nombre,
      empleadoId: perfil.empleado_id,
      preferencias: perfil.preferencias || { pantallaInicio: 'dashboard', moneda: '$', idioma: 'es' },
      ultimoAcceso: perfil.ultimo_acceso,
      accesos: perfil.accesos,
      empresas: esAdministradorGlobal ? 'todas' : empresasIds,
      rol: rolEnEmpresa(empresasIds[0]),
      _rolPorEmpresa: accesos,
    };
  };

  useEffect(() => {
    supabase.auth.getSession().then(async ({ data: { session } }) => {
      if (session?.user) {
        const s = await cargarSesionDesdeAuth(session.user);
        if (s) setSesion(s);
      }
      setAuthLoading(false);
    });
    const { data: listener } = supabase.auth.onAuthStateChange(async (_event, session) => {
      if (!session?.user) { setSesion(null); return; }
      const s = await cargarSesionDesdeAuth(session.user);
      if (s) setSesion(s);
    });
    return () => listener.subscription.unsubscribe();
  }, []);

  const intentarLogin = async () => {
    setLoginError('');
    let email = loginEmail.trim();
    if (!email.includes('@')) {
      // No parece un correo -> se asume que escribió su "usuario" corto; se resuelve a su correo real.
      const { data: emailEncontrado, error: errBusqueda } = await supabase.rpc('email_por_username', { p_username: email.toLowerCase() });
      if (errBusqueda || !emailEncontrado) { setLoginError('Usuario o correo incorrectos.'); return; }
      email = emailEncontrado;
    }
    const { data, error } = await supabase.auth.signInWithPassword({ email, password: loginClave });
    if (error) { setLoginError('Usuario/correo o clave incorrectos.'); return; }
    const ahora = new Date().toISOString();
    await supabase.from('perfiles').update({ ultimo_acceso: ahora, accesos: (data.user.user_metadata?.accesos || 0) + 1 }).eq('id', data.user.id);
    setLoginClave('');
    const s = await cargarSesionDesdeAuth(data.user);
    if (s) {
      const defaultPorRol = s.rol === 'Tecnico' ? 'agenda' : s.rol === 'Operaciones' ? 'ordenes' : 'dashboard';
      const preferida = s.preferencias?.pantallaInicio;
      setTab(preferida && tieneAcceso(s.rol, preferida) ? preferida : defaultPorRol);
    }
  };

  const intentarRegistro = async () => {
    setLoginError('');
    if (!loginNombre.trim()) { setLoginError('Escribe tu nombre completo.'); return; }
    const { data, error } = await supabase.auth.signUp({
      email: loginEmail.trim(), password: loginClave, options: { data: { nombre: loginNombre.trim() } },
    });
    if (error) { setLoginError(error.message === 'User already registered' ? 'Ese correo ya está registrado — inicia sesión.' : 'No se pudo crear la cuenta. ' + error.message); return; }
    if (data.user && !data.session) {
      setLoginError('');
      alert('Cuenta creada. Revisa tu correo para confirmar la cuenta antes de iniciar sesión (según la configuración del proyecto).');
      setAuthMode('login');
    }
    // Nota: esta cuenta nace SIN acceso a ninguna empresa todavía — un Administrador debe asignárselo
    // desde Usuarios y roles (o, para el primer usuario del sistema, se hace directo en Supabase).
  };

  const cerrarSesion = async () => {
    await supabase.auth.signOut();
    setSesion(null);
  };

  if (authLoading) {
    return (
      <div style={{ fontFamily: 'Inter, sans-serif', backgroundColor: C.navy, minHeight: '100vh', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
        <div style={{ color: '#fff', fontSize: 13 }}>Cargando…</div>
      </div>
    );
  }

  if (!sesion) {
    return (
      <div style={{ fontFamily: 'Inter, sans-serif', backgroundColor: C.navy, minHeight: '100vh', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 16 }}>
        <style>{`@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600;700&display=swap'); * { box-sizing: border-box; }`}</style>
        <div className="rounded-2xl p-8" style={{ backgroundColor: '#fff', width: 380, maxWidth: '100%', boxShadow: '0 24px 60px rgba(0,0,0,0.35)' }}>
          <div className="flex items-center gap-3 mb-6">
            <svg width="42" height="42" viewBox="0 0 100 100">
              <path d="M 78 6 C 42 14, 20 34, 16 58" stroke={C.navy} strokeWidth="17" fill="none" strokeLinecap="round" />
              <path d="M 16 58 C 18 74, 30 84, 44 87" stroke={C.navy} strokeWidth="13" fill="none" strokeLinecap="round" />
              <line x1="40" y1="86" x2="72" y2="76" stroke={C.navy} strokeWidth="4.5" strokeLinecap="round" />
              <line x1="42" y1="90" x2="78" y2="84" stroke={C.navy} strokeWidth="4.5" strokeLinecap="round" />
              <line x1="42" y1="94" x2="76" y2="93" stroke={C.navy} strokeWidth="4.5" strokeLinecap="round" />
              <circle cx="76" cy="76" r="3.4" fill={C.cyanDark} />
              <circle cx="82" cy="84" r="3.4" fill={C.cyanDark} />
              <circle cx="80" cy="93" r="3.4" fill={C.cyanDark} />
            </svg>
            <div>
              <div style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 700, fontSize: 17, color: C.slate }}>BUILDERTECH SRL</div>
              <div style={{ fontSize: 11, color: C.muted, fontStyle: 'italic' }}>Noi lo facciamo meglio</div>
            </div>
          </div>
          <div style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 600, fontSize: 15, color: C.slate, marginBottom: 14 }}>
            {authMode === 'login' ? 'Iniciar sesión' : 'Crear cuenta'}
          </div>
          {authMode === 'signup' && (
            <div className="mb-3"><Field label="Nombre completo"><input style={{ ...inputStyle, width: '100%' }} value={loginNombre} onChange={(e) => setLoginNombre(e.target.value)} autoFocus /></Field></div>
          )}
          <Field label={authMode === 'login' ? 'Usuario o correo' : 'Correo'}><input type={authMode === 'login' ? 'text' : 'email'} style={{ ...inputStyle, width: '100%' }} value={loginEmail} onChange={(e) => setLoginEmail(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && (authMode === 'login' ? intentarLogin() : intentarRegistro())} autoFocus={authMode === 'login'} /></Field>
          <div className="mt-3"><Field label="Clave"><input type="password" style={{ ...inputStyle, width: '100%' }} value={loginClave} onChange={(e) => setLoginClave(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && (authMode === 'login' ? intentarLogin() : intentarRegistro())} /></Field></div>
          {loginError && <div style={{ fontSize: 12, color: C.coral, marginTop: 8, fontWeight: 600 }}>{loginError}</div>}
          <button onClick={authMode === 'login' ? intentarLogin : intentarRegistro} className="w-full mt-5 py-2.5 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.cyanDark }}>
            {authMode === 'login' ? 'Entrar' : 'Crear cuenta'}
          </button>
          <button onClick={() => { setAuthMode(authMode === 'login' ? 'signup' : 'login'); setLoginError(''); }} className="w-full mt-3 py-2 rounded-lg text-sm font-semibold" style={{ backgroundColor: '#fff', color: C.muted }}>
            {authMode === 'login' ? '¿No tienes cuenta? Créala aquí' : '¿Ya tienes cuenta? Inicia sesión'}
          </button>
          {authMode === 'signup' && (
            <div style={{ fontSize: 11, color: C.muted, marginTop: 12, textAlign: 'center' }}>
              Tu cuenta se crea sin acceso a ninguna empresa todavía — un Administrador debe asignártelo desde Usuarios y roles.
            </div>
          )}
        </div>
      </div>
    );
  }

  return (
    <div style={{ fontFamily: 'Inter, sans-serif', backgroundColor: C.bg, minHeight: '100vh', display: 'flex' }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@500;600&display=swap');
        * { box-sizing: border-box; }
        table { border-collapse: collapse; width: 100%; }
        th { text-align: left; font-size: 11px; text-transform: uppercase; letter-spacing: 0.04em; color: ${C.muted}; font-weight: 600; padding: 10px 12px; border-bottom: 1px solid ${C.border}; }
        td { padding: 11px 12px; font-size: 13px; border-bottom: 1px solid ${C.border}; color: ${C.slate}; }
        tr:last-child td { border-bottom: none; }
        @media print {
          aside, .no-print { display: none !important; }
          body, html { background: #fff !important; }
          .flex-1 { padding: 0 !important; max-width: 100% !important; }
        }
        @media (max-width: 767px) {
          /* Las grillas de tarjetas (KPIs, resúmenes) pasan a una sola columna en celular — excepto las marcadas .keep-grid-cols (como el calendario) */
          [class*="grid-cols-"]:not(.keep-grid-cols) { grid-template-columns: 1fr !important; }
          /* Las tablas anchas se desplazan horizontalmente en vez de recortarse */
          .rounded-xl.overflow-hidden { overflow-x: auto !important; -webkit-overflow-scrolling: touch; }
          table { min-width: 640px; }
        }
        select, input { font-family: 'Inter', sans-serif; }
        button { cursor: pointer; font-family: 'Inter', sans-serif; }
      `}</style>

      {/* Sidebar (escritorio) */}
      <div className="no-print hidden md:block" style={{ width: 232, backgroundColor: C.navy, flexShrink: 0, minHeight: '100vh', padding: '20px 14px' }}>
        <div className="flex items-center gap-2.5 px-2 mb-8">
          {empresaActual?.nombre === 'BUILDERTECH SRL' ? (
            <svg width="34" height="34" viewBox="0 0 100 100" style={{ flexShrink: 0 }}>
              {/* Logo BuilderTech: cable curvo con fibras */}
              <path d="M 78 6 C 42 14, 20 34, 16 58" stroke="#fff" strokeWidth="17" fill="none" strokeLinecap="round" />
              <path d="M 16 58 C 18 74, 30 84, 44 87" stroke="#fff" strokeWidth="13" fill="none" strokeLinecap="round" />
              <line x1="40" y1="86" x2="72" y2="76" stroke="#fff" strokeWidth="4.5" strokeLinecap="round" />
              <line x1="42" y1="90" x2="78" y2="84" stroke="#fff" strokeWidth="4.5" strokeLinecap="round" />
              <line x1="42" y1="94" x2="76" y2="93" stroke="#fff" strokeWidth="4.5" strokeLinecap="round" />
              <circle cx="76" cy="76" r="3.4" fill={C.cyan} />
              <circle cx="82" cy="84" r="3.4" fill={C.cyan} />
              <circle cx="80" cy="93" r="3.4" fill={C.cyan} />
            </svg>
          ) : (
            <div style={{ width: 34, height: 34, borderRadius: 9, backgroundColor: C.cyan, color: C.navy, display: 'flex', alignItems: 'center', justifyContent: 'center', fontFamily: 'Space Grotesk, sans-serif', fontWeight: 700, fontSize: 16, flexShrink: 0 }}>
              {(empresaActual?.nombre || '?').trim()[0].toUpperCase()}
            </div>
          )}
          <div>
            <div style={{ color: '#fff', fontFamily: 'Space Grotesk, sans-serif', fontWeight: 700, fontSize: 14.5, letterSpacing: '0.02em' }}>{(empresaActual?.nombre || 'BUILDERTECH SRL').toUpperCase()}</div>
            <div style={{ color: '#7C93A8', fontSize: 10, fontStyle: 'italic' }}>{empresaActual?.eslogan || empresaActual?.giro || ''}</div>
          </div>
        </div>
        <div className="flex flex-col gap-0.5">
          {nav.map((n) => {
            if (n.children) {
              const isOpen = expandedGroup === n.key;
              const childActive = n.children.some((c) => c.key === tab);
              return (
                <div key={n.key}>
                  <button onClick={() => setExpandedGroup(isOpen ? null : n.key)} className="flex items-center justify-between w-full px-3 py-2.5 rounded-lg text-left transition-colors"
                    style={{ backgroundColor: childActive ? C.navyLight : 'transparent', color: childActive ? C.cyan : '#93A5B8', fontSize: 13, fontWeight: childActive ? 600 : 500, borderLeft: childActive ? `2px solid ${C.cyan}` : '2px solid transparent' }}>
                    <span className="flex items-center gap-2.5"><n.icon size={16} />{n.label}</span>
                    <ChevronDown size={14} style={{ transform: isOpen ? 'rotate(180deg)' : 'none', transition: 'transform 0.15s' }} />
                  </button>
                  {isOpen && (
                    <div className="flex flex-col gap-0.5 mt-0.5">
                      {n.children.map((c) => (
                        <button key={c.key} onClick={() => setTab(c.key)} className="flex items-center px-3 py-2 rounded-lg text-left transition-colors"
                          style={{ marginLeft: 22, backgroundColor: tab === c.key ? C.navyLight : 'transparent', color: tab === c.key ? C.cyan : '#7C93A8', fontSize: 12.5, fontWeight: tab === c.key ? 600 : 500 }}>
                          {c.label}
                        </button>
                      ))}
                    </div>
                  )}
                </div>
              );
            }
            const Icon = n.icon;
            const active = tab === n.key;
            return (
              <button key={n.key} onClick={() => setTab(n.key)} className="flex items-center gap-2.5 px-3 py-2.5 rounded-lg text-left transition-colors"
                style={{ backgroundColor: active ? C.navyLight : 'transparent', color: active ? C.cyan : '#93A5B8', fontSize: 13, fontWeight: active ? 600 : 500, borderLeft: active ? `2px solid ${C.cyan}` : '2px solid transparent' }}>
                <Icon size={16} />{n.label}
              </button>
            );
          })}
        </div>
      </div>

      {/* Columna derecha: barra móvil + contenido, apilados verticalmente junto al sidebar */}
      <div className="flex-1" style={{ display: 'flex', flexDirection: 'column', minWidth: 0 }}>
      {/* Barra superior móvil (< md): hamburguesa + marca + drawer de navegación */}
      <div className="no-print flex md:hidden items-center justify-between px-3" style={{ backgroundColor: C.navy, height: 56, position: 'sticky', top: 0, zIndex: 45 }}>
        <button onClick={() => setMobileNavOpen(true)} className="flex items-center justify-center" style={{ width: 36, height: 36 }}>
          <div className="flex flex-col gap-1"><span style={{ width: 20, height: 2, backgroundColor: '#fff', borderRadius: 2 }} /><span style={{ width: 20, height: 2, backgroundColor: '#fff', borderRadius: 2 }} /><span style={{ width: 20, height: 2, backgroundColor: '#fff', borderRadius: 2 }} /></div>
        </button>
        <div style={{ color: '#fff', fontFamily: 'Space Grotesk, sans-serif', fontWeight: 700, fontSize: 13, letterSpacing: '0.02em' }}>{(empresaActual?.nombre || 'BUILDERTECH SRL').toUpperCase()}</div>
        <div className="flex items-center gap-2">
          <button onClick={() => setNotifPanelOpen(!notifPanelOpen)} style={{ position: 'relative', width: 34, height: 34, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
            <Bell size={18} color="#fff" />
            {notificaciones.filter((n) => n.usuarioId === sesion.id && !n.leida).length > 0 && (
              <span style={{ position: 'absolute', top: 2, right: 2, backgroundColor: C.coral, borderRadius: '50%', width: 9, height: 9 }} />
            )}
          </button>
          <button onClick={() => setUserMenuOpen(!userMenuOpen)} style={{ width: 36, height: 36, borderRadius: '50%', backgroundColor: C.cyan, color: C.navy, display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 700, fontSize: 12 }}>
            {sesion.nombre.split(' ').map((p) => p[0]).slice(0, 2).join('')}
          </button>
        </div>
      </div>
      {mobileNavOpen && (
        <div className="no-print md:hidden" style={{ position: 'fixed', inset: 0, zIndex: 70, backgroundColor: 'rgba(11,22,34,0.55)' }} onClick={() => setMobileNavOpen(false)}>
          <div onClick={(e) => e.stopPropagation()} style={{ width: 270, maxWidth: '85vw', height: '100%', backgroundColor: C.navy, padding: '20px 14px', overflowY: 'auto' }}>
            <div className="flex items-center justify-between mb-6 px-2">
              <div style={{ color: '#fff', fontFamily: 'Space Grotesk, sans-serif', fontWeight: 700, fontSize: 14.5 }}>{(empresaActual?.nombre || 'BUILDERTECH SRL').toUpperCase()}</div>
              <button onClick={() => setMobileNavOpen(false)}><X size={20} color="#fff" /></button>
            </div>
            <div className="flex flex-col gap-0.5">
              {nav.map((n) => {
                if (n.children) {
                  const isOpen = expandedGroup === n.key;
                  const childActive = n.children.some((c) => c.key === tab);
                  return (
                    <div key={n.key}>
                      <button onClick={() => setExpandedGroup(isOpen ? null : n.key)} className="flex items-center justify-between w-full px-3 py-2.5 rounded-lg text-left"
                        style={{ backgroundColor: childActive ? C.navyLight : 'transparent', color: childActive ? C.cyan : '#93A5B8', fontSize: 13.5, fontWeight: childActive ? 600 : 500 }}>
                        <span className="flex items-center gap-2.5"><n.icon size={16} />{n.label}</span>
                        <ChevronDown size={14} style={{ transform: isOpen ? 'rotate(180deg)' : 'none' }} />
                      </button>
                      {isOpen && (
                        <div className="flex flex-col gap-0.5 mt-0.5">
                          {n.children.map((c) => (
                            <button key={c.key} onClick={() => { setTab(c.key); setMobileNavOpen(false); }} className="flex items-center px-3 py-2.5 rounded-lg text-left"
                              style={{ marginLeft: 22, backgroundColor: tab === c.key ? C.navyLight : 'transparent', color: tab === c.key ? C.cyan : '#7C93A8', fontSize: 13, fontWeight: tab === c.key ? 600 : 500 }}>
                              {c.label}
                            </button>
                          ))}
                        </div>
                      )}
                    </div>
                  );
                }
                const Icon = n.icon;
                const active = tab === n.key;
                return (
                  <button key={n.key} onClick={() => { setTab(n.key); setMobileNavOpen(false); }} className="flex items-center gap-2.5 px-3 py-2.5 rounded-lg text-left"
                    style={{ backgroundColor: active ? C.navyLight : 'transparent', color: active ? C.cyan : '#93A5B8', fontSize: 13.5, fontWeight: active ? 600 : 500 }}>
                    <Icon size={16} />{n.label}
                  </button>
                );
              })}
            </div>
          </div>
        </div>
      )}

      {/* Main */}
      <div className="flex-1 p-4 md:p-8" style={{ maxWidth: 1220, minWidth: 0, overflowX: 'hidden' }}>
        {confirmState && (
          <div style={{ position: 'fixed', inset: 0, backgroundColor: 'rgba(11,22,34,0.45)', zIndex: 60, display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 16 }}
            onClick={() => setConfirmState(null)}>
            <div onClick={(e) => e.stopPropagation()} className="rounded-2xl p-5" style={{ backgroundColor: '#fff', width: 400, maxWidth: '100%', boxShadow: '0 20px 50px rgba(0,0,0,0.25)' }}>
              <div style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 700, fontSize: 15, color: C.slate, marginBottom: 10 }}>Confirmar acción</div>
              <div style={{ fontSize: 13, color: C.slate, whiteSpace: 'pre-wrap', lineHeight: 1.5 }}>{confirmState.mensaje}</div>
              <div className="flex gap-2 justify-end mt-5">
                <button onClick={() => setConfirmState(null)} className="px-4 py-2 rounded-lg text-sm font-semibold" style={{ backgroundColor: '#EEF1F4', color: C.slate }}>Cancelar</button>
                <button onClick={() => { const fn = confirmState.onConfirm; setConfirmState(null); fn(); }} className="px-4 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.coral }}>Confirmar</button>
              </div>
            </div>
          </div>
        )}
        <div className="flex justify-end items-center gap-2.5 mb-4 no-print" style={{ position: 'relative' }}>
          {(() => {
            const misNotifs = notificaciones.filter((n) => n.usuarioId === sesion.id);
            const noLeidas = misNotifs.filter((n) => !n.leida).length;
            return (
              <div style={{ position: 'relative' }}>
                <button onClick={() => setNotifPanelOpen(!notifPanelOpen)} className="hidden md:flex items-center justify-center rounded-xl" style={{ width: 42, height: 42, backgroundColor: C.card, border: `1px solid ${C.border}`, position: 'relative' }}>
                  <Bell size={18} color={C.slate} />
                  {noLeidas > 0 && (
                    <span style={{ position: 'absolute', top: -4, right: -4, backgroundColor: C.coral, color: '#fff', fontSize: 10, fontWeight: 700, borderRadius: '50%', width: 18, height: 18, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                      {noLeidas > 9 ? '9+' : noLeidas}
                    </span>
                  )}
                </button>
                {notifPanelOpen && (
                  <div className="rounded-xl overflow-hidden" style={{ position: 'absolute', top: 50, right: 0, backgroundColor: '#fff', border: `1px solid ${C.border}`, boxShadow: '0 12px 32px rgba(0,0,0,0.14)', width: 340, zIndex: 40, maxHeight: 420, overflowY: 'auto' }}>
                    <div className="px-4 py-3 flex items-center justify-between" style={{ backgroundColor: C.navy }}>
                      <span style={{ color: '#fff', fontWeight: 700, fontSize: 13 }}>Notificaciones</span>
                      <div className="flex items-center gap-3">
                        {noLeidas > 0 && (
                          <button onClick={() => { setNotificaciones((prev) => prev.map((n) => n.usuarioId === sesion.id ? { ...n, leida: true } : n)); supabase.from('notificaciones').update({ leida: true }).eq('usuario_id', sesion.id).then(() => {}); }} style={{ color: C.cyan, fontSize: 11, fontWeight: 600 }}>Marcar todas leídas</button>
                        )}
                        {misNotifs.some((n) => n.leida) && (
                          <button onClick={() => { setNotificaciones((prev) => prev.filter((n) => !(n.usuarioId === sesion.id && n.leida))); supabase.from('notificaciones').delete().eq('usuario_id', sesion.id).eq('leida', true).then(() => {}); }} style={{ color: '#9DB2C4', fontSize: 11, fontWeight: 600 }}>Limpiar leídas</button>
                        )}
                      </div>
                    </div>
                    {misNotifs.length === 0 && <div style={{ padding: 20, textAlign: 'center', color: C.muted, fontSize: 12.5 }}>Sin notificaciones todavía.</div>}
                    {misNotifs.map((n) => {
                      const iconoTipo = { orden_asignada: '📋', orden_por_validar: '⏳', orden_completada: '✅' }[n.tipo] || '🔔';
                      const minAgo = Math.max(0, Math.floor((new Date() - new Date(n.fecha)) / 60000));
                      const cuandoTxt = minAgo < 1 ? 'ahora mismo' : minAgo < 60 ? `hace ${minAgo} min` : minAgo < 1440 ? `hace ${Math.floor(minAgo / 60)} h` : `hace ${Math.floor(minAgo / 1440)} d`;
                      return (
                        <button key={n.id} onClick={() => {
                          setNotificaciones((prev) => prev.map((x) => x.id === n.id ? { ...x, leida: true } : x));
                          supabase.from('notificaciones').update({ leida: true }).eq('id', n.id).then(() => {});
                          if (n.ordenId) { setTab(sesion.rol === 'Tecnico' ? 'agenda' : 'ordenes'); }
                          setNotifPanelOpen(false);
                        }} className="w-full text-left px-4 py-3 flex gap-2.5" style={{ backgroundColor: n.leida ? '#fff' : '#F0FBF9', borderBottom: `1px solid ${C.border}` }}>
                          <span style={{ fontSize: 16 }}>{iconoTipo}</span>
                          <div>
                            <div style={{ fontSize: 12.5, fontWeight: n.leida ? 500 : 700, color: C.slate }}>{n.titulo}</div>
                            <div style={{ fontSize: 11.5, color: C.muted, marginTop: 2 }}>{n.mensaje}</div>
                            <div style={{ fontSize: 10, color: C.muted, marginTop: 3 }}>{cuandoTxt}</div>
                          </div>
                          {!n.leida && <span style={{ width: 7, height: 7, borderRadius: '50%', backgroundColor: C.cyanDark, flexShrink: 0, marginTop: 4 }} />}
                        </button>
                      );
                    })}
                  </div>
                )}
              </div>
            );
          })()}
          <button onClick={() => setUserMenuOpen(!userMenuOpen)} className="hidden md:flex items-center gap-2.5 px-3 py-2 rounded-xl" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
            <div style={{ width: 30, height: 30, borderRadius: '50%', backgroundColor: C.navy, color: '#fff', display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 700, fontSize: 13 }}>
              {sesion.nombre.split(' ').map((p) => p[0]).slice(0, 2).join('')}
            </div>
            <div style={{ textAlign: 'left' }}>
              <div style={{ fontSize: 12.5, fontWeight: 700, color: C.slate }}>{sesion.nombre}</div>
              <div style={{ fontSize: 10.5, color: C.cyanDark, fontWeight: 600 }}>● {sesion.rol}</div>
            </div>
            <ChevronDown size={14} color={C.muted} />
          </button>
          {userMenuOpen && (
            <div className="rounded-xl overflow-hidden" style={{ position: 'absolute', top: 52, right: 0, backgroundColor: '#fff', border: `1px solid ${C.border}`, boxShadow: '0 12px 32px rgba(0,0,0,0.14)', width: 220, zIndex: 40 }}>
              <div className="px-4 py-3" style={{ backgroundColor: C.navy }}>
                <div style={{ color: '#fff', fontWeight: 700, fontSize: 13 }}>{sesion.nombre}</div>
                <div style={{ color: '#9DB2C4', fontSize: 11 }}>@{sesion.usuario} · {sesion.rol}</div>
                <div style={{ color: C.cyan, fontSize: 10.5, marginTop: 3 }}>🏢 {empresaActual?.nombre}</div>
              </div>
              {empresasDeSesion().length > 1 && (
                <div>
                  <button onClick={() => setEmpresaMenuOpen(!empresaMenuOpen)} className="w-full text-left px-4 py-2.5 text-sm font-semibold" style={{ color: C.slate }}>
                    {t('Cambiar empresa')} {empresaMenuOpen ? '▴' : '▾'}
                  </button>
                  {empresaMenuOpen && empresasDeSesion().map((e) => (
                    <button key={e.id} onClick={() => cambiarEmpresa(e.id)} className="w-full text-left px-6 py-2 text-sm" style={{ color: e.id === empresaActiva ? C.cyanDark : C.slate, fontWeight: e.id === empresaActiva ? 700 : 400, backgroundColor: e.id === empresaActiva ? '#E4F8F1' : 'transparent' }}>
                      {e.id === empresaActiva ? '✓ ' : ''}{e.nombre}
                      <div style={{ fontSize: 10, color: C.muted }}>{e.giro}</div>
                    </button>
                  ))}
                </div>
              )}
              {sesion.rol === 'Administrador' && (
                <button onClick={() => { setTab('empresas'); setUserMenuOpen(false); }} className="w-full text-left px-4 py-2.5 text-sm" style={{ color: C.slate, borderTop: `1px solid ${C.border}` }}>{t('Empresas')}</button>
              )}
              {sesion.rol === 'Administrador' && (
                <button onClick={() => { setTab('usuarios'); setUserMenuOpen(false); }} className="w-full text-left px-4 py-2.5 text-sm" style={{ color: C.slate }}>{t('Usuarios y roles')}</button>
              )}
              <button onClick={() => { setCuentaNombreEdit(sesion.nombre); setCuentaUsernameEdit(sesion.username || ''); setCuentaClaveNueva(''); setCuentaClaveConfirm(''); setCuentaMsg(''); setTab('cuenta'); setUserMenuOpen(false); }} className="w-full text-left px-4 py-2.5 text-sm" style={{ color: C.slate, borderTop: `1px solid ${C.border}` }}>{t('Cuenta')}</button>
              <button onClick={() => { setTab('preferencias'); setUserMenuOpen(false); }} className="w-full text-left px-4 py-2.5 text-sm" style={{ color: C.slate }}>{t('Preferencias')}</button>
              <button onClick={() => { setTab('formato_correos'); setUserMenuOpen(false); }} className="w-full text-left px-4 py-2.5 text-sm" style={{ color: C.slate }}>{t('Formato de correos')}</button>
              <button onClick={() => { cerrarSesion(); setUserMenuOpen(false); }} className="w-full text-left px-4 py-2.5 text-sm font-semibold" style={{ color: C.coral, borderTop: `1px solid ${C.border}` }}>{t('Cerrar sesión')}</button>
            </div>
          )}
        </div>
        {sesion && esSoloLectura(sesion.rol, tab) && (
          <div className="rounded-xl p-3 mb-4" style={{ backgroundColor: '#EAF2FB', border: '1px solid #3B82C455', fontSize: 12.5, color: '#1D5A96', fontWeight: 600 }}>
            👁 Estás en modo solo lectura en este módulo: puedes consultar pero no modificar.
          </div>
        )}
        {/* DASHBOARD */}
        {tab === 'dashboard' && (
          <div>
            <h1 style={{ fontFamily: 'Space Grotesk, sans-serif', fontSize: 22, fontWeight: 700, color: C.slate }}>{t('Indicadores KPI')}</h1>
            <p style={{ color: C.muted, fontSize: 13, marginTop: 2 }}>Resumen financiero, operativo y de inventario en tiempo real.</p>
            <div className="grid grid-cols-4 gap-4 mt-6">
              <KPICard label="Ingresos facturados" value={fmt(revenue)} sub={`${completedOrders.length} orden(es) reconocida(s)`} accent={C.cyanDark} />
              <KPICard label="Costos totales" value={fmt(costs)} sub="Materiales + mano de obra" />
              <KPICard label="Utilidad neta" value={fmt(profit)} sub={`Margen ${revenue ? ((profit / revenue) * 100).toFixed(1) : '0.0'}%`} accent={profit >= 0 ? C.cyanDark : C.coral} />
              <KPICard label="Cuentas por cobrar" value={fmt(receivable)} sub="Facturas sin pagar" accent={C.amber} />
              <KPICard label="Pagos pendientes a técnicos" value={fmt(techPending)} sub="Mano de obra sin liquidar" accent={C.amber} />
              <KPICard label="Órdenes activas" value={inProgressCount} sub="Pendientes o en proceso" />
              <KPICard label="Valor de inventario" value={fmt(inventoryValue)} sub="Costo promedio × stock" accent={C.cyanDark} />
              <KPICard label="Proyectos activos" value={activeProjects} sub={`${projects.length} proyectos en total`} />
              <KPICard label="Cuentas por pagar" value={fmt(cxpPendiente)} sub="A proveedores" accent={C.amber} />
              <KPICard label="Activos fijos (valor en libros)" value={fmt(valorActivosLibros)} sub={`${activos.length} activos registrados`} />
            </div>

            {lowStock.length > 0 && (
              <div className="mt-5 rounded-xl p-3.5 flex items-start gap-2.5" style={{ backgroundColor: '#FFF4E0', border: `1px solid ${C.amber}55` }}>
                <AlertTriangle size={16} color={C.amber} style={{ marginTop: 2, flexShrink: 0 }} />
                <div style={{ fontSize: 12.5, color: '#8A5A00' }}><strong>Inventario bajo:</strong> {lowStock.map((m) => `${m.nombre} (${m.stock} ${m.unidad})`).join(', ')}.</div>
              </div>
            )}

            <div className="mt-7 rounded-xl overflow-hidden" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
              <div className="px-4 pt-4 pb-1" style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 600, fontSize: 14, color: C.slate }}>Órdenes recientes</div>
              <table>
                <thead><tr><th>Folio</th><th>Proyecto</th><th>Cliente</th><th>Técnico</th><th>Fecha</th><th>Estado</th><th>Precio</th></tr></thead>
                <tbody>
                  {[...orders].slice(-6).reverse().map((o) => (
                    <tr key={o.id}>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 12 }}>{o.id}</td>
                      <td>{o.proyectoId ? projectName(o.proyectoId) : '—'}</td>
                      <td>{clientName(o.clienteId)}</td>
                      <td>{techName(o.tecnicoId)}</td>
                      <td>{o.fecha}</td>
                      <td><SignalBadge estado={o.estado} /></td>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(o.precioCliente)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* KPI: INDICES FINANCIEROS */}
        {tab === 'kpi_financieros' && (
          <div>
            <h1 style={{ fontFamily: 'Space Grotesk, sans-serif', fontSize: 22, fontWeight: 700, color: C.slate }}>{t('Índices financieros')}</h1>
            <p style={{ color: C.muted, fontSize: 13, marginTop: 2 }}>Calculados en tiempo real del Balance General y el Estado de Resultados.</p>
            <div className="flex gap-3 items-end mt-4">
              <Field label="Del"><input type="date" style={inputStyle} value={reportFrom} onChange={(e) => setReportFrom(e.target.value)} /></Field>
              <Field label="Al"><input type="date" style={inputStyle} value={reportTo} onChange={(e) => setReportTo(e.target.value)} /></Field>
            </div>
            <div className="grid grid-cols-3 gap-4 mt-5">
              {[
                ['Liquidez corriente', kpiFin.liquidez, 'x', 'Activo corriente / Pasivo corriente. Sano: mayor a 1.5', kpiFin.liquidez === null ? C.muted : kpiFin.liquidez >= 1.5 ? C.cyanDark : kpiFin.liquidez >= 1 ? C.amber : C.coral],
                ['Capital de trabajo', kpiFin.capitalTrabajo, '$', 'Activo corriente − Pasivo corriente', kpiFin.capitalTrabajo >= 0 ? C.cyanDark : C.coral],
                ['Endeudamiento', kpiFin.endeudamiento, '%', 'Pasivo total / Activo total. Sano: menor a 60%', kpiFin.endeudamiento === null ? C.muted : kpiFin.endeudamiento <= 0.6 ? C.cyanDark : C.coral],
                ['Margen operativo', kpiFin.margenOperativo, '%', 'Utilidad operativa / Ingresos del periodo', kpiFin.margenOperativo === null ? C.muted : kpiFin.margenOperativo >= 0 ? C.cyanDark : C.coral],
                ['Margen neto', kpiFin.margenNeto, '%', 'Utilidad neta / Ingresos del periodo', kpiFin.margenNeto === null ? C.muted : kpiFin.margenNeto >= 0 ? C.cyanDark : C.coral],
                ['Rentabilidad del capital (ROE)', kpiFin.roe, '%', 'Utilidad neta / Capital total', kpiFin.roe === null ? C.muted : kpiFin.roe >= 0 ? C.cyanDark : C.coral],
              ].map(([label, val, unidad, desc, color], i) => (
                <div key={i} className="rounded-xl p-4" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
                  <div style={{ fontSize: 11, fontWeight: 700, color: C.muted, textTransform: 'uppercase' }}>{label}</div>
                  <div style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 22, fontWeight: 700, color, marginTop: 6 }}>
                    {val === null ? '—' : unidad === '$' ? fmt(val) : unidad === '%' ? (val * 100).toFixed(1) + '%' : val.toFixed(2) + 'x'}
                  </div>
                  <div style={{ fontSize: 11, color: C.muted, marginTop: 6 }}>{desc}</div>
                </div>
              ))}
            </div>
            <div className="grid grid-cols-2 gap-5 mt-5">
              <BarChart title="Márgenes del periodo" valueFmt={(v) => (v * 100).toFixed(1) + '%'}
                colorFn={(d) => d.value >= 0 ? C.cyanDark : C.coral}
                data={[
                  { label: 'Margen operativo', value: kpiFin.margenOperativo || 0 },
                  { label: 'Margen neto', value: kpiFin.margenNeto || 0 },
                  { label: 'ROE', value: kpiFin.roe || 0 },
                ]} />
              <DonutChart title="Composición del activo"
                data={[
                  { label: 'Activo corriente', value: Math.max(0, activoCorriente), color: C.cyanDark },
                  { label: 'Activo no corriente', value: Math.max(0, totalActivo - activoCorriente), color: C.navy },
                ]} />
            </div>
          </div>
        )}

        {/* KPI: GESTION DE ORDENES */}
        {tab === 'kpi_ordenes' && (
          <div>
            <h1 style={{ fontFamily: 'Space Grotesk, sans-serif', fontSize: 22, fontWeight: 700, color: C.slate }}>{t('Gestión de órdenes')}</h1>
            <p style={{ color: C.muted, fontSize: 13, marginTop: 2 }}>El equivalente a "gestión de ventas" en una empresa de servicios: tus órdenes de instalación.</p>
            <div className="flex gap-3 items-end mt-4">
              <Field label="Del"><input type="date" style={inputStyle} value={reportFrom} onChange={(e) => setReportFrom(e.target.value)} /></Field>
              <Field label="Al"><input type="date" style={inputStyle} value={reportTo} onChange={(e) => setReportTo(e.target.value)} /></Field>
            </div>
            <div className="grid grid-cols-4 gap-4 mt-5">
              {[
                ['Órdenes en el periodo', kpiOrd.total, null],
                ['Completadas', kpiOrd.completadas, C.cyanDark],
                ['Canceladas', kpiOrd.canceladas, kpiOrd.canceladas > 0 ? C.coral : null],
                ['Ingreso promedio por orden', fmt(kpiOrd.ticketPromedio), C.cyanDark],
              ].map(([label, val, color], i) => (
                <div key={i} className="rounded-xl p-4" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
                  <div style={{ fontSize: 11, fontWeight: 700, color: C.muted, textTransform: 'uppercase' }}>{label}</div>
                  <div style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 22, fontWeight: 700, color: color || C.slate, marginTop: 6 }}>{val}</div>
                </div>
              ))}
            </div>
            <div className="rounded-xl p-4 mt-4" style={{ backgroundColor: C.card, border: `1px solid ${C.border}`, maxWidth: 420 }}>
              <div style={{ fontSize: 11, fontWeight: 700, color: C.muted, textTransform: 'uppercase' }}>Utilidad promedio por orden completada</div>
              <div style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 22, fontWeight: 700, color: kpiOrd.utilidadPromedio >= 0 ? C.cyanDark : C.coral, marginTop: 6 }}>{fmt(kpiOrd.utilidadPromedio)}</div>
              <div style={{ fontSize: 11, color: C.muted, marginTop: 4 }}>Precio al cliente menos materiales y mano de obra.</div>
            </div>
            <div className="grid grid-cols-2 gap-5 mt-5">
              <BarChart title="Utilidad por centro de costo" valueFmt={fmt}
                colorFn={(d) => d.value >= 0 ? C.cyanDark : C.coral}
                data={centros.filter((c) => c.codigo !== 'CC-ADM').map((c) => ({ label: c.nombre, value: centroCostoMetrics(c.id, ordersRango.filter((o) => o.estado === 'completada')).utilidad }))} />
              <DonutChart title="Órdenes del periodo por estado"
                data={[
                  { label: 'Completadas', value: ordersRango.filter((o) => o.estado === 'completada').length, color: C.cyanDark },
                  { label: 'En proceso', value: ordersRango.filter((o) => o.estado === 'en_proceso').length, color: '#3B82C4' },
                  { label: 'Pendientes', value: ordersRango.filter((o) => o.estado === 'pendiente').length, color: C.amber },
                  { label: 'Canceladas', value: ordersRango.filter((o) => o.estado === 'cancelada').length, color: C.coral },
                ].filter((d) => d.value > 0)} />
            </div>
            <div className="rounded-xl overflow-hidden mt-5" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
              <div className="px-4 pt-4 pb-1" style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 600, fontSize: 14, color: C.slate }}>Desempeño por centro de costo</div>
              <table>
                <thead><tr><th>Centro</th><th>Órdenes</th><th>Ingresos</th><th>Costo materiales</th><th>Costo MO</th><th>Utilidad</th></tr></thead>
                <tbody>
                  {centros.filter((c) => c.codigo !== 'CC-ADM').map((c) => {
                    const m = centroCostoMetrics(c.id, ordersRango.filter((o) => o.estado === 'completada'));
                    return (
                      <tr key={c.id}>
                        <td>{c.nombre}</td>
                        <td>{m.ordenes}</td>
                        <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(m.ingresos)}</td>
                        <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(m.cMateriales)}</td>
                        <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(m.cManoObra)}</td>
                        <td style={{ fontFamily: 'JetBrains Mono, monospace', color: m.utilidad >= 0 ? C.cyanDark : C.coral, fontWeight: 600 }}>{fmt(m.utilidad)}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* KPI: RENDIMIENTO DE TECNICOS */}
        {tab === 'kpi_tecnicos' && (
          <div>
            <h1 style={{ fontFamily: 'Space Grotesk, sans-serif', fontSize: 22, fontWeight: 700, color: C.slate }}>{t('Rendimiento de técnicos')}</h1>
            <p style={{ color: C.muted, fontSize: 13, marginTop: 2 }}>Quién genera más instalaciones, ingresos y utilidad en el periodo.</p>
            <div className="flex gap-3 items-end mt-4">
              <Field label="Del"><input type="date" style={inputStyle} value={reportFrom} onChange={(e) => setReportFrom(e.target.value)} /></Field>
              <Field label="Al"><input type="date" style={inputStyle} value={reportTo} onChange={(e) => setReportTo(e.target.value)} /></Field>
            </div>
            <div className="grid grid-cols-2 gap-5 mt-5">
              <BarChart title="Ingresos generados por técnico" valueFmt={fmt}
                data={[...kpiTecnicos].sort((a, b) => b.ingresos - a.ingresos).map((t) => ({ label: t.nombre, value: t.ingresos }))} />
              <BarChart title="Órdenes completadas por técnico" valueFmt={(v) => String(v)}
                colorFn={() => C.cyan}
                data={[...kpiTecnicos].sort((a, b) => b.completadas - a.completadas).map((t) => ({ label: t.nombre, value: t.completadas }))} />
            </div>
            <div className="rounded-xl overflow-hidden mt-5" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
              <table>
                <thead><tr><th>Técnico</th><th>Órdenes asignadas</th><th>Completadas</th><th>Ingresos generados</th><th>Utilidad generada</th><th>Saldo pendiente de pago</th></tr></thead>
                <tbody>
                  {kpiTecnicos.map((t) => (
                    <tr key={t.id}>
                      <td>{t.nombre}</td>
                      <td>{t.ordenes}</td>
                      <td>{t.completadas}</td>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(t.ingresos)}</td>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace', color: t.utilidad >= 0 ? C.cyanDark : C.coral }}>{fmt(t.utilidad)}</td>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace', color: C.amber }}>{fmt(orders.filter((o) => o.tecnicoId === t.id && o.estadoPagoTecnico === 'pendiente' && o.estado !== 'cancelada').reduce((s, o) => s + Number(o.costoManoObra), 0))}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* KPI: INVENTARIO Y ROTACION */}
        {tab === 'kpi_inventario' && (
          <div>
            <h1 style={{ fontFamily: 'Space Grotesk, sans-serif', fontSize: 22, fontWeight: 700, color: C.slate }}>{t('Inventario y rotación')}</h1>
            <p style={{ color: C.muted, fontSize: 13, marginTop: 2 }}>Rotación de materiales de instalación y alertas de stock.</p>
            <div className="flex gap-3 items-end mt-4">
              <Field label="Del"><input type="date" style={inputStyle} value={reportFrom} onChange={(e) => setReportFrom(e.target.value)} /></Field>
              <Field label="Al"><input type="date" style={inputStyle} value={reportTo} onChange={(e) => setReportTo(e.target.value)} /></Field>
            </div>
            <div className="grid grid-cols-3 gap-4 mt-5">
              {[
                ['Valor del inventario', fmt(kpiInv.valorInventario), C.slate, 'Al costo promedio ponderado actual'],
                ['Consumo del periodo', fmt(kpiInv.consumoPeriodo), C.slate, 'Salidas de kardex valorizadas'],
                ['Rotación', kpiInv.rotacion === null ? '—' : kpiInv.rotacion.toFixed(2) + 'x', kpiInv.rotacion !== null && kpiInv.rotacion >= 0.5 ? C.cyanDark : C.amber, 'Consumo / Valor de inventario. Más alto = inventario más eficiente'],
              ].map(([label, val, color, desc], i) => (
                <div key={i} className="rounded-xl p-4" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
                  <div style={{ fontSize: 11, fontWeight: 700, color: C.muted, textTransform: 'uppercase' }}>{label}</div>
                  <div style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 22, fontWeight: 700, color, marginTop: 6 }}>{val}</div>
                  <div style={{ fontSize: 11, color: C.muted, marginTop: 6 }}>{desc}</div>
                </div>
              ))}
            </div>
            <div className="mt-5">
              <BarChart title="Valor en inventario por material (top 8)" valueFmt={fmt}
                colorFn={(d) => d.stockBajo ? C.coral : C.cyanDark}
                data={[...materials].sort((a, b) => (b.stock * b.costoUnitario) - (a.stock * a.costoUnitario)).slice(0, 8).map((m) => ({ label: m.nombre, value: m.stock * m.costoUnitario, stockBajo: m.stock < 10 }))} />
            </div>
            <div className="grid grid-cols-2 gap-5 mt-5">
              <div className="rounded-xl overflow-hidden" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
                <div className="px-4 pt-4 pb-1" style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 600, fontSize: 14, color: C.slate }}>Top 5 materiales más consumidos</div>
                <table>
                  <thead><tr><th>Material</th><th>Consumido</th><th>Stock actual</th></tr></thead>
                  <tbody>
                    {kpiInv.topConsumidos.map((m) => (
                      <tr key={m.id}><td>{m.nombre}</td><td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{m.consumido} {m.unidad}</td><td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{m.stock}</td></tr>
                    ))}
                    {kpiInv.topConsumidos.length === 0 && <tr><td colSpan={3} style={{ color: C.muted, textAlign: 'center', padding: 16 }}>Sin consumos en el periodo.</td></tr>}
                  </tbody>
                </table>
              </div>
              <div className="rounded-xl overflow-hidden" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
                <div className="px-4 pt-4 pb-1" style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 600, fontSize: 14, color: C.slate }}>Alerta: stock bajo (≤ 10 unidades)</div>
                <table>
                  <thead><tr><th>Material</th><th>Stock</th><th>Valor</th></tr></thead>
                  <tbody>
                    {kpiInv.stockBajo.map((m) => (
                      <tr key={m.id}><td>{m.nombre}</td><td style={{ fontFamily: 'JetBrains Mono, monospace', color: C.coral, fontWeight: 700 }}>{m.stock}</td><td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(m.stock * m.costoUnitario)}</td></tr>
                    ))}
                    {kpiInv.stockBajo.length === 0 && <tr><td colSpan={3} style={{ color: C.muted, textAlign: 'center', padding: 16 }}>Sin materiales en nivel crítico. ✓</td></tr>}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {/* KPI: CARTERA Y PAGOS */}
        {tab === 'kpi_cartera' && (
          <div>
            <h1 style={{ fontFamily: 'Space Grotesk, sans-serif', fontSize: 22, fontWeight: 700, color: C.slate }}>{t('Cartera y pagos')}</h1>
            <p style={{ color: C.muted, fontSize: 13, marginTop: 2 }}>Antigüedad de lo que te deben (CxC) y lo que debes (CxP).</p>
            <div className="grid grid-cols-3 gap-4 mt-5">
              {[
                ['Por cobrar a clientes', fmt(cxcTotal), C.amber, `${cxcPendientes.length} orden(es) completadas sin cobrar`],
                ['Por pagar a proveedores', fmt(cxpPendiente), C.slate, `${cxpVencidas.length} factura(s) ya vencidas`],
                ['Nómina y MO pendiente', fmt(nominaPendiente + orders.filter((o) => o.estadoPagoTecnico === 'pendiente' && o.estado !== 'cancelada').reduce((s, o) => s + Number(o.costoManoObra), 0)), C.slate, 'Empleados + técnicos'],
              ].map(([label, val, color, desc], i) => (
                <div key={i} className="rounded-xl p-4" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
                  <div style={{ fontSize: 11, fontWeight: 700, color: C.muted, textTransform: 'uppercase' }}>{label}</div>
                  <div style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 22, fontWeight: 700, color, marginTop: 6 }}>{val}</div>
                  <div style={{ fontSize: 11, color: C.muted, marginTop: 6 }}>{desc}</div>
                </div>
              ))}
            </div>
            <div className="mt-5">
              <DonutChart title="Composición de pendientes"
                data={[
                  { label: 'Por cobrar a clientes', value: cxcTotal, color: C.amber },
                  { label: 'Por pagar a proveedores', value: cxpPendiente, color: C.slate },
                  { label: 'Nómina y MO pendiente', value: nominaPendiente + orders.filter((o) => o.estadoPagoTecnico === 'pendiente' && o.estado !== 'cancelada').reduce((s, o) => s + Number(o.costoManoObra), 0), color: '#3B82C4' },
                ].filter((d) => d.value > 0)} />
            </div>
            <div className="grid grid-cols-2 gap-5 mt-5">
              <div className="rounded-xl p-5" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
                <div style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 600, fontSize: 14, color: C.slate, marginBottom: 10 }}>Antigüedad de cartera (CxC)</div>
                {Object.entries(agingCxc).map(([rango, monto]) => (
                  <div key={rango} className="flex items-center gap-3 py-1.5">
                    <span style={{ fontSize: 12.5, color: C.slate, width: 110 }}>{rango}</span>
                    <div style={{ flex: 1, height: 10, backgroundColor: '#EEF1F4', borderRadius: 5, overflow: 'hidden' }}>
                      <div style={{ width: cxcTotal > 0 ? `${(monto / cxcTotal) * 100}%` : 0, height: '100%', backgroundColor: rango === 'Más de 60 días' ? C.coral : rango === '31-60 días' ? C.amber : C.cyanDark }} />
                    </div>
                    <span style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 12.5, width: 100, textAlign: 'right' }}>{fmt(monto)}</span>
                  </div>
                ))}
              </div>
              <div className="rounded-xl overflow-hidden" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
                <div className="px-4 pt-4 pb-1" style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 600, fontSize: 14, color: C.slate }}>CxP próximas a vencer</div>
                <table>
                  <thead><tr><th>Concepto</th><th>Vence</th><th>Monto</th></tr></thead>
                  <tbody>
                    {cxpVencidas.map((c) => (
                      <tr key={c.id}><td style={{ fontSize: 12.5 }}>{c.concepto}</td><td style={{ color: C.coral, fontWeight: 700, fontSize: 12 }}>{c.vencimiento} (vencida)</td><td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(c.monto)}</td></tr>
                    ))}
                    {cxpPorVencer.map((c) => (
                      <tr key={c.id}><td style={{ fontSize: 12.5 }}>{c.concepto}</td><td style={{ fontSize: 12 }}>{c.vencimiento}</td><td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(c.monto)}</td></tr>
                    ))}
                    {cxpVencidas.length === 0 && cxpPorVencer.length === 0 && <tr><td colSpan={3} style={{ color: C.muted, textAlign: 'center', padding: 16 }}>Sin cuentas por pagar pendientes. ✓</td></tr>}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {/* ORDENES */}
        {tab === 'ordenes' && (
          <div>
            <div className="flex items-center justify-between">
              <div>
                <h1 style={{ fontFamily: 'Space Grotesk, sans-serif', fontSize: 22, fontWeight: 700, color: C.slate }}>{t('Órdenes de trabajo')}</h1>
                <p style={{ color: C.muted, fontSize: 13, marginTop: 2 }}>{orders.length} órdenes registradas · {activeOrders.length} activas</p>
              </div>
              {sesion.rol !== 'Tecnico' && (
                <div className="flex gap-2">
                  <label className="flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-white text-sm font-semibold cursor-pointer" style={{ backgroundColor: C.cyanDark }}>
                    <Download size={15} style={{ transform: 'rotate(180deg)' }} /> {t('Importar Excel')}
                    <input type="file" accept=".xlsx,.xls,.csv" style={{ display: 'none' }} onChange={(e) => { if (e.target.files[0]) importarOrdenesExcel(e.target.files[0]); e.target.value = ''; }} />
                  </label>
                  <button onClick={() => setShowOrderForm(!showOrderForm)} className="flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.navy }}>
                    {showOrderForm ? <X size={15} /> : <Plus size={15} />} {showOrderForm ? t('Cerrar') : t('Nueva orden')}
                  </button>
                </div>
              )}
            </div>

            {importPreview.length > 0 && (
              <div className="rounded-xl overflow-hidden mt-5" style={{ backgroundColor: C.card, border: `2px solid ${C.cyanDark}` }}>
                <div className="px-4 pt-4 pb-2 flex items-center justify-between flex-wrap gap-2">
                  <div>
                    <span style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 600, fontSize: 14, color: C.slate }}>
                      Vista previa — {importPreview.filter((p) => !p.duplicado).length} nueva(s) · {importPreview.filter((p) => p.duplicado).length} ya existente(s) · {importPreview.filter((p) => p.incluir).length} seleccionadas
                    </span>
                    <div style={{ fontSize: 11.5, color: C.muted, marginTop: 2 }}>Las órdenes ya importadas vienen desmarcadas automáticamente. Asigna técnico, ajusta fecha/hora y define precio y MO antes de crear.</div>
                  </div>
                  <div className="flex gap-2">
                    <button onClick={crearOrdenesImportadas} className="px-4 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.cyanDark }}>Crear {importPreview.filter((p) => p.incluir).length} orden(es)</button>
                    <button onClick={() => setImportPreview([])} className="px-3 py-2 rounded-lg text-sm font-semibold" style={{ backgroundColor: '#EEF1F4', color: C.slate }}>Cancelar</button>
                  </div>
                </div>
                <div style={{ maxHeight: 420, overflowY: 'auto' }}>
                  <table>
                    <thead><tr><th></th><th>Referencia</th><th>Dirección</th><th>Fecha</th><th>Hora</th><th>Técnico</th><th>Centro</th><th>Precio ($)</th><th>MO ($)</th></tr></thead>
                    <tbody>
                      {importPreview.map((p) => (
                        <tr key={p.idx} style={{ opacity: p.incluir ? 1 : 0.45, backgroundColor: p.duplicado ? '#FFF7E8' : 'transparent' }}>
                          <td><input type="checkbox" checked={p.incluir} onChange={(e) => updateImportRow(p.idx, 'incluir', e.target.checked)} style={{ width: 15, height: 15, accentColor: C.cyanDark }} /></td>
                          <td style={{ fontSize: 11.5 }}>
                            {p.referencia}{p.wr && <div style={{ color: C.muted }}>WR {p.wr}</div>}{p.notas && <div style={{ color: C.amber, fontSize: 10.5 }}>{p.notas}</div>}
                            {p.duplicado && <div className="text-xs font-bold px-1.5 py-0.5 rounded mt-1" style={{ backgroundColor: '#FDEBD0', color: '#8A5A00', display: 'inline-block' }}>⚠ {p.duplicado}</div>}
                          </td>
                          <td style={{ fontSize: 12 }}>{p.direccion}</td>
                          <td><input type="date" value={p.fecha} onChange={(e) => updateImportRow(p.idx, 'fecha', e.target.value)} style={{ ...inputStyle, fontSize: 11.5, padding: '3px 5px', width: 125 }} /></td>
                          <td><input type="time" value={p.hora} onChange={(e) => updateImportRow(p.idx, 'hora', e.target.value)} style={{ ...inputStyle, fontSize: 11.5, padding: '3px 5px', width: 80 }} /></td>
                          <td>
                            <select value={p.tecnicoId} onChange={(e) => updateImportRow(p.idx, 'tecnicoId', e.target.value)} style={{ ...inputStyle, fontSize: 11.5, padding: '3px 5px' }}>
                              {techs.map((t) => <option key={t.id} value={t.id}>{t.nombre}</option>)}
                            </select>
                          </td>
                          <td>
                            <select value={p.centroCostoId} onChange={(e) => updateImportRow(p.idx, 'centroCostoId', e.target.value)} style={{ ...inputStyle, fontSize: 11.5, padding: '3px 5px' }}>
                              {centros.map((c) => <option key={c.id} value={c.id}>{c.nombre}</option>)}
                            </select>
                          </td>
                          <td><input type="number" value={p.precioCliente} onChange={(e) => updateImportRow(p.idx, 'precioCliente', e.target.value)} placeholder="0" style={{ ...inputStyle, fontSize: 11.5, padding: '3px 5px', width: 75 }} /></td>
                          <td><input type="number" value={p.costoManoObra} onChange={(e) => updateImportRow(p.idx, 'costoManoObra', e.target.value)} placeholder="0" style={{ ...inputStyle, fontSize: 11.5, padding: '3px 5px', width: 75 }} /></td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {showOrderForm && (
              <OrderForm clients={clients.filter((c) => c.activo !== false)} techs={techs} materials={materials} projects={projects} centros={centros} catalogoPartidas={catalogoPartidas} tiposServicio={tiposServicio}
                onSubmit={(data) => { addOrderWithKardex(data); setShowOrderForm(false); }} />
            )}

            <div className="flex gap-1.5 mt-5 mb-3 items-center flex-wrap">
              {['todas', 'por_validar', 'pendiente', 'en_proceso', 'completada', 'cancelada'].map((f) => (
                <button key={f} onClick={() => setOrderFilter(f)} className="px-3 py-1.5 rounded-full text-xs font-semibold"
                  style={{ backgroundColor: orderFilter === f ? C.navy : '#fff', color: orderFilter === f ? '#fff' : C.muted, border: `1px solid ${C.border}` }}>
                  {f === 'todas' ? 'Todas' : ESTADOS[f].label}{f === 'por_validar' && orders.filter((o) => o.estado === 'por_validar').length > 0 ? ` (${orders.filter((o) => o.estado === 'por_validar').length})` : ''}
                </button>
              ))}
              <div className="flex items-center gap-1.5 ml-auto" style={{ position: 'relative' }}>
                <input value={ordersSearch} onChange={(e) => setOrdersSearch(e.target.value)} placeholder="Buscar folio, cliente, técnico, dirección..."
                  style={{ ...inputStyle, width: 240, paddingLeft: 30 }} />
                <span style={{ position: 'absolute', left: 9, color: C.muted, fontSize: 13, pointerEvents: 'none' }}>🔍</span>
                {ordersSearch && (
                  <button onClick={() => setOrdersSearch('')} style={{ position: 'absolute', right: 8, color: C.muted }}><X size={13} /></button>
                )}
              </div>
            </div>
            {ordersSearch && (
              <div style={{ fontSize: 11.5, color: C.muted, marginTop: -6, marginBottom: 8 }}>
                {visibleOrders.length} resultado{visibleOrders.length === 1 ? '' : 's'} para "{ordersSearch}"
              </div>
            )}

            <div className="rounded-xl overflow-hidden" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
              <table>
                <thead><tr><th>{t('Folio')}</th><th>{t('Proyecto')}</th><th>{t('Centro de costo')}</th><th>{t('Cliente')}</th><th>{t('Técnico')}</th><th>Servicio</th><th>{t('Fecha')}</th><th>Costo</th><th>Precio</th><th>Utilidad</th><th>{t('Estado')}</th><th></th></tr></thead>
                <tbody>
                  {ordersPag.paginaActual.map((o) => (
                    <tr key={o.id}>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 12 }}>{o.id}{o.referencia && <div style={{ fontSize: 10, color: C.muted, fontFamily: 'Inter, sans-serif' }}>{o.referencia}</div>}</td>
                      <td style={{ fontSize: 12 }}>{o.proyectoId ? projectName(o.proyectoId) : '—'}</td>
                      <td style={{ fontSize: 12 }}>{centroCostoName(o.centroCostoId)}</td>
                      <td>{clientName(o.clienteId)}</td>
                      <td>{techName(o.tecnicoId)}</td>
                      <td>{o.tipoServicio}</td>
                      <td>{o.fecha}{o.hora && <div style={{ fontSize: 10.5, color: C.muted }}>{o.hora}</div>}</td>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmtSeg(orderCost(o), 'ordenes')}</td>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmtSeg(o.precioCliente, 'ordenes')}</td>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace', color: orderProfit(o) >= 0 ? C.cyanDark : C.coral }}>{fmtSeg(orderProfit(o), 'ordenes')}</td>
                      <td>
                        {o.estado === 'por_validar' ? (
                          <div className="flex items-center gap-1.5">
                            <span className="text-xs font-semibold px-2 py-0.5 rounded-full" style={{ backgroundColor: '#F1EBFB', color: '#8E6FCB' }}>Por validar</span>
                            <button onClick={() => validarOrden(o.id, true)} title="Validar" className="p-1 rounded" style={{ backgroundColor: '#E4F8F1' }}><Check size={13} color={C.cyanDark} /></button>
                            <button onClick={() => validarOrden(o.id, false)} title="Rechazar" className="p-1 rounded" style={{ backgroundColor: '#FDE8E8' }}><X size={13} color={C.coral} /></button>
                          </div>
                        ) : (
                          <select value={o.estado} onChange={(e) => changeOrderEstado(o.id, e.target.value)} style={{ ...inputStyle, padding: '4px 6px', fontSize: 12 }}>
                            {Object.keys(ESTADOS).filter((k) => k !== 'por_validar').map((k) => <option key={k} value={k}>{ESTADOS[k].label}</option>)}
                          </select>
                        )}
                      </td>
                      <td>
                        <button onClick={() => deleteOrderWithReversal(o.id)}
                          title={o.estado === 'cancelada' ? 'Eliminar definitivamente' : 'Solo se pueden eliminar órdenes canceladas — cancélala primero'}
                          style={{ opacity: o.estado === 'cancelada' ? 1 : 0.35, cursor: o.estado === 'cancelada' ? 'pointer' : 'not-allowed' }}>
                          <Trash2 size={14} color={C.coral} />
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
                {visibleOrders.length > 0 && (
                  <tfoot>
                    <tr>
                      <td colSpan={7} style={{ fontWeight: 700 }}>Total ({visibleOrders.length} orden{visibleOrders.length === 1 ? '' : 'es'})</td>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace', fontWeight: 700 }}>{fmtSeg(visibleOrders.reduce((s, o) => s + orderCost(o), 0), 'ordenes')}</td>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace', fontWeight: 700 }}>{fmtSeg(visibleOrders.reduce((s, o) => s + Number(o.precioCliente || 0), 0), 'ordenes')}</td>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace', fontWeight: 700, color: visibleOrders.reduce((s, o) => s + orderProfit(o), 0) >= 0 ? C.cyanDark : C.coral }}>{fmtSeg(visibleOrders.reduce((s, o) => s + orderProfit(o), 0), 'ordenes')}</td>
                      <td></td><td></td>
                    </tr>
                  </tfoot>
                )}
              </table>
              <Paginador p={ordersPag} />
            </div>
          </div>
        )}

        {/* PROYECTOS */}
        {tab === 'proyectos' && (
          <div>
            <div className="flex items-center justify-between">
              <div>
                <h1 style={{ fontFamily: 'Space Grotesk, sans-serif', fontSize: 22, fontWeight: 700, color: C.slate }}>{t('Proyectos')}</h1>
                <p style={{ color: C.muted, fontSize: 13, marginTop: 2 }}>Agrupa varias órdenes bajo un mismo contrato y compara costo real contra presupuesto.</p>
              </div>
              <button onClick={() => setShowProjectForm(!showProjectForm)} className="flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.navy }}>
                {showProjectForm ? <X size={15} /> : <Plus size={15} />} {showProjectForm ? 'Cerrar' : 'Nuevo proyecto'}
              </button>
            </div>

            {showProjectForm && (
              <div className="rounded-xl p-4 mt-4 flex flex-wrap gap-3 items-end" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
                <ProjectFormFields clients={clients} onSubmit={async (data) => {
                  const { data: nuevo, error } = await supabase.from('proyectos').insert({
                    empresa_id: empresaActiva, nombre: data.nombre, cliente_id: data.clienteId, fecha_inicio: data.fechaInicio,
                    presupuesto: Number(data.presupuesto) || 0, estado: 'activo',
                  }).select().single();
                  if (error) { alert('No se pudo guardar el proyecto: ' + error.message); return; }
                  setProjects((p) => [...p, { id: nuevo.id, estado: 'activo', ...data }]);
                  setShowProjectForm(false);
                }} />
              </div>
            )}

            <div className="grid grid-cols-2 gap-3 mt-5">
              {projects.map((p) => {
                const m = projectMetrics(p);
                const pct = Math.min(m.avance * 100, 999);
                const barColor = m.avance > 1 ? C.coral : m.avance > 0.85 ? C.amber : C.cyanDark;
                return (
                  <div key={p.id} className="rounded-xl p-4" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
                    <div className="flex items-center justify-between">
                      <div style={{ fontWeight: 600, fontSize: 14, color: C.slate }}>{p.nombre}</div>
                      <select value={p.estado} onChange={(e) => { setProjects((prev) => prev.map((x) => x.id === p.id ? { ...x, estado: e.target.value } : x)); supabase.from('proyectos').update({ estado: e.target.value }).eq('id', p.id).then(({ error }) => { if (error) console.error(error); }); }}
                        style={{ ...inputStyle, padding: '3px 6px', fontSize: 11, color: ESTADOS_PROYECTO[p.estado].color }}>
                        {Object.keys(ESTADOS_PROYECTO).map((k) => <option key={k} value={k}>{ESTADOS_PROYECTO[k].label}</option>)}
                      </select>
                    </div>
                    <div style={{ color: C.muted, fontSize: 12.5, marginTop: 2 }}>{clientName(p.clienteId)} · inicio {p.fechaInicio} · {m.linked.length} órdenes</div>

                    <div className="mt-3">
                      <div className="flex justify-between text-xs mb-1" style={{ fontFamily: 'JetBrains Mono, monospace' }}>
                        <span style={{ color: C.muted }}>Presupuesto {fmt(p.presupuesto)}</span>
                        <span style={{ color: barColor, fontWeight: 600 }}>{pct.toFixed(0)}% usado</span>
                      </div>
                      <div style={{ height: 6, borderRadius: 3, backgroundColor: C.bg, overflow: 'hidden' }}>
                        <div style={{ height: '100%', width: `${Math.min(pct, 100)}%`, backgroundColor: barColor, borderRadius: 3 }} />
                      </div>
                    </div>

                    <div className="flex gap-4 mt-3 text-xs" style={{ fontFamily: 'JetBrains Mono, monospace' }}>
                      <span style={{ color: C.slate }}>Gastado: {fmt(m.gastado)}</span>
                      <span style={{ color: C.cyanDark }}>Facturado: {fmt(m.ingresos)}</span>
                      <span style={{ color: m.utilidad >= 0 ? C.cyanDark : C.coral }}>Utilidad: {fmt(m.utilidad)}</span>
                    </div>
                  </div>
                );
              })}
              {projects.length === 0 && <div style={{ color: C.muted, fontSize: 13 }}>Aún no hay proyectos registrados.</div>}
            </div>
          </div>
        )}

        {/* CLIENTES */}
        {tab === 'clientes' && (
          <div>
            <div className="flex items-center justify-between">
              <h1 style={{ fontFamily: 'Space Grotesk, sans-serif', fontSize: 22, fontWeight: 700, color: C.slate }}>{t('Clientes')}</h1>
              <button onClick={() => setShowClientForm(!showClientForm)} className="flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.navy }}>
                {showClientForm ? <X size={15} /> : <Plus size={15} />} {showClientForm ? 'Cerrar' : 'Nuevo cliente'}
              </button>
            </div>
            {showClientForm && (
              <SimpleForm fields={[{ key: 'nombre', label: 'Nombre / Razón social' }, { key: 'telefono', label: 'Teléfono' }, { key: 'direccion', label: 'Dirección' }]}
                onSubmit={async (data) => {
                  const { data: nuevo, error } = await supabase.from('clientes').insert({ empresa_id: empresaActiva, nombre: data.nombre, telefono: data.telefono, direccion: data.direccion }).select().single();
                  if (error) { alert('No se pudo guardar el cliente: ' + error.message); return; }
                  setClients((p) => [...p, { id: nuevo.id, ...data, activo: true }]);
                  setShowClientForm(false);
                }} />
            )}
            <div className="grid grid-cols-2 gap-3 mt-5">
              {clients.map((c) => {
                const total = orders.filter((o) => o.clienteId === c.id && o.estado === 'completada').reduce((s, o) => s + Number(o.precioCliente), 0);
                return (
                  <div key={c.id} className="rounded-xl p-4" style={{ backgroundColor: C.card, border: `1px solid ${C.border}`, opacity: c.activo === false ? 0.55 : 1 }}>
                    <div className="flex items-center justify-between">
                      <input value={c.nombre} onChange={(ev) => updateCliente(c.id, 'nombre', ev.target.value)} style={{ ...inputStyle, fontWeight: 600, fontSize: 13.5, padding: '3px 6px', width: 200 }} />
                      <div className="flex items-center gap-2">
                        {c.activo === false && <span className="text-xs font-bold px-2 py-0.5 rounded-full" style={{ backgroundColor: '#FDE8E8', color: C.coral }}>INACTIVO</span>}
                        <button onClick={() => toggleClienteActivo(c.id)} className="text-xs font-semibold px-2 py-1 rounded-lg" style={{ backgroundColor: c.activo === false ? '#E4F8F1' : '#FDE8E8', color: c.activo === false ? C.cyanDark : C.coral }}>
                          {c.activo === false ? 'Reactivar' : 'Dar de baja'}
                        </button>
                      </div>
                    </div>
                    <div className="flex items-center gap-1.5 mt-2" style={{ color: C.muted, fontSize: 12.5 }}><MapPin size={12} /><input value={c.direccion || ''} onChange={(ev) => updateCliente(c.id, 'direccion', ev.target.value)} style={{ ...inputStyle, fontSize: 11.5, padding: '3px 6px', width: 220 }} /></div>
                    <div className="flex items-center gap-1.5 mt-1" style={{ color: C.muted, fontSize: 12.5 }}><Phone size={12} /><input value={c.telefono || ''} onChange={(ev) => updateCliente(c.id, 'telefono', ev.target.value)} style={{ ...inputStyle, fontSize: 11.5, padding: '3px 6px', width: 140 }} /></div>
                    <div className="mt-2.5 text-xs" style={{ color: C.cyanDark, fontFamily: 'JetBrains Mono, monospace', fontWeight: 600 }}>Facturado: {fmt(total)}</div>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {/* TECNICOS (vista de los empleados con rol Técnico) */}
        {tab === 'tecnicos' && (
          <div>
            <div className="flex items-center justify-between">
              <div>
                <h1 style={{ fontFamily: 'Space Grotesk, sans-serif', fontSize: 22, fontWeight: 700, color: C.slate }}>{t('Técnicos')}</h1>
                <p style={{ color: C.muted, fontSize: 13, marginTop: 2 }}>Los técnicos son empleados con rol "Técnico" — se dan de alta y editan en Recursos humanos → Empleados.</p>
              </div>
              <button onClick={() => setTab('rh')} className="flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.navy }}>
                <Plus size={15} /> Alta en Empleados
              </button>
            </div>
            <div className="grid grid-cols-2 gap-3 mt-5">
              {techs.map((t) => {
                const done = orders.filter((o) => o.tecnicoId === t.id && o.estado === 'completada');
                const ganado = done.reduce((s, o) => s + Number(o.costoManoObra), 0);
                const pagado = done.filter((o) => o.estadoPagoTecnico === 'pagado').reduce((s, o) => s + Number(o.costoManoObra), 0);
                return (
                  <div key={t.id} className="rounded-xl p-4" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
                    <div style={{ fontWeight: 600, fontSize: 14, color: C.slate }}>{t.nombre}</div>
                    <div style={{ color: C.muted, fontSize: 12.5, marginTop: 2 }}>{t.puesto} · {t.telefono}</div>
                    <div className="flex gap-4 mt-2.5 text-xs" style={{ fontFamily: 'JetBrains Mono, monospace' }}>
                      <span style={{ color: C.slate }}>Ganado: {fmt(ganado)}</span>
                      <span style={{ color: C.cyanDark }}>Pagado: {fmt(pagado)}</span>
                      <span style={{ color: C.amber }}>Debe: {fmt(ganado - pagado)}</span>
                    </div>
                  </div>
                );
              })}
              {techs.length === 0 && <div style={{ color: C.muted, fontSize: 13 }}>Aún no hay empleados con rol Técnico. Da de alta uno en Recursos humanos → Empleados.</div>}
            </div>
          </div>
        )}

        {/* INVENTARIO */}
        {tab === 'inventario' && (
          <div>
            <div className="flex items-center justify-between">
              <h1 style={{ fontFamily: 'Space Grotesk, sans-serif', fontSize: 22, fontWeight: 700, color: C.slate }}>{t('Inventario de materiales')}</h1>
              <button onClick={() => setShowMatForm(!showMatForm)} className="flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.navy }}>
                {showMatForm ? <X size={15} /> : <Plus size={15} />} {showMatForm ? 'Cerrar' : 'Nuevo material'}
              </button>
            </div>
            {showMatForm && (
              <SimpleForm fields={[{ key: 'nombre', label: 'Nombre' }, { key: 'unidad', label: 'Unidad (pza, rollo...)' }, { key: 'costoUnitario', label: 'Costo unitario ($)', type: 'number' }, { key: 'stock', label: 'Stock inicial', type: 'number' }]}
                onSubmit={(data) => { addMaterialWithKardex(data); setShowMatForm(false); }} />
            )}
            <div className="rounded-xl overflow-hidden mt-5" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
              <table>
                <thead><tr><th>Material</th><th>Unidad</th><th>Costo promedio</th><th>Stock</th><th>Valor en inventario</th><th>Estado</th></tr></thead>
                <tbody>
                  {materials.map((m) => (
                    <tr key={m.id}>
                      <td>{m.nombre}</td>
                      <td>{m.unidad}</td>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(m.costoUnitario)}</td>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{m.stock}</td>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(m.stock * m.costoUnitario)}</td>
                      <td>{m.stock < 10
                        ? <span className="text-xs font-semibold px-2 py-0.5 rounded-full" style={{ backgroundColor: '#FFF4E0', color: C.amber }}>Stock bajo</span>
                        : <span className="text-xs font-semibold px-2 py-0.5 rounded-full" style={{ backgroundColor: '#E4F8F1', color: C.cyanDark }}>Suficiente</span>}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* EQUIPOS EN COMODATO — dentro del módulo de Inventario (no consumibles, con serial). Específico de telecom/fibra óptica. */}
        {tab === 'inventario' && empresaActual?.giro === 'Servicios de fibra óptica' && (
          <div className="mt-8 pt-6" style={{ borderTop: `1px solid ${C.border}` }}>
            <div className="flex items-center justify-between">
              <div>
                <h2 style={{ fontFamily: 'Space Grotesk, sans-serif', fontSize: 17, fontWeight: 700, color: C.slate }}>Equipos en comodato (ONT / CPE)</h2>
                <p style={{ color: C.muted, fontSize: 12.5, marginTop: 2 }}>Equipos propiedad del operador entregados al cliente en comodato — se rastrean por número de serie, no por costeo promedio.</p>
              </div>
              <button onClick={() => setShowEquipoForm(!showEquipoForm)} className="flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.navy }}>
                {showEquipoForm ? <X size={15} /> : <Plus size={15} />} {showEquipoForm ? 'Cerrar' : 'Nuevo equipo'}
              </button>
            </div>

            {showEquipoForm && <EquipoComodatoForm onSubmit={(data) => { addEquipoComodato(data); setShowEquipoForm(false); }} />}

            <div className="grid grid-cols-3 gap-3 mt-4">
              {[
                ['En almacén', equiposComodato.filter((e) => e.estado === 'en_almacen').length, C.cyanDark],
                ['Asignados a clientes', equiposComodato.filter((e) => e.estado === 'asignado').length, C.amber],
                ['De baja', equiposComodato.filter((e) => e.estado === 'baja').length, C.coral],
              ].map(([label, val, color], i) => <KPICard key={i} label={label} value={String(val)} accent={color} />)}
            </div>

            <div className="rounded-xl overflow-hidden mt-4" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
              <table>
                <thead><tr><th>Serial</th><th>Tipo</th><th>Modelo</th><th>Sector</th><th>Estado</th><th>Cliente asignado</th><th>Fecha</th><th></th></tr></thead>
                <tbody>
                  {equiposComodato.map((e) => (
                    <tr key={e.id} style={{ opacity: e.estado === 'baja' ? 0.5 : 1 }}>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 12 }}>{e.serial}</td>
                      <td><span className="text-xs font-semibold px-2 py-0.5 rounded-full" style={{ backgroundColor: e.tipo === 'ONT' ? '#E4F8F1' : '#EEF1F4', color: e.tipo === 'ONT' ? C.cyanDark : C.slate }}>{e.tipo}</span></td>
                      <td style={{ fontSize: 12.5 }}>{e.modelo}</td>
                      <td style={{ fontSize: 12, color: C.muted }}>{e.sector || '—'}</td>
                      <td>
                        {e.estado === 'en_almacen' && <span className="text-xs font-semibold px-2 py-0.5 rounded-full" style={{ backgroundColor: '#E4F8F1', color: C.cyanDark }}>En almacén</span>}
                        {e.estado === 'asignado' && <span className="text-xs font-semibold px-2 py-0.5 rounded-full" style={{ backgroundColor: '#FFF4E0', color: C.amber }}>Asignado</span>}
                        {e.estado === 'baja' && <span className="text-xs font-semibold px-2 py-0.5 rounded-full" style={{ backgroundColor: '#FDE8E8', color: C.coral }}>Baja</span>}
                      </td>
                      <td style={{ fontSize: 12.5 }}>{e.clienteId ? clientName(e.clienteId) : '—'}</td>
                      <td style={{ fontSize: 11.5, color: C.muted }}>{e.fechaAsignacion || '—'}</td>
                      <td className="flex gap-1.5">
                        {e.estado === 'en_almacen' && (
                          <select onChange={(ev) => { if (ev.target.value) asignarEquipoComodato(e.id, ev.target.value); ev.target.value = ''; }} defaultValue="" style={{ ...inputStyle, fontSize: 11, padding: '3px 5px' }}>
                            <option value="" disabled>Asignar a...</option>
                            {clients.filter((c) => c.activo !== false).map((c) => <option key={c.id} value={c.id}>{c.nombre}</option>)}
                          </select>
                        )}
                        {e.estado === 'asignado' && <button onClick={() => devolverEquipoComodato(e.id)} className="text-xs font-semibold px-2 py-1 rounded-lg" style={{ backgroundColor: '#EEF1F4', color: C.slate }}>Devolver</button>}
                        {e.estado !== 'baja' && <button onClick={() => bajaEquipoComodato(e.id)} title="Dar de baja"><Trash2 size={13} color={C.coral} /></button>}
                      </td>
                    </tr>
                  ))}
                  {equiposComodato.length === 0 && <tr><td colSpan={8} style={{ color: C.muted, textAlign: 'center', padding: 20 }}>Sin equipos registrados.</td></tr>}
                </tbody>
              </table>
            </div>
          </div>
        )}
        {tab === 'kardex' && (
          <div>
            <h1 style={{ fontFamily: 'Space Grotesk, sans-serif', fontSize: 22, fontWeight: 700, color: C.slate }}>{t('Kardex de inventario')}</h1>
            <p style={{ color: C.muted, fontSize: 13, marginTop: 2 }}>Costeo promedio ponderado. Las órdenes generan salidas automáticas; registra aquí compras (entradas) o consumos manuales (salidas) con su centro de costo.</p>

            <div className="flex flex-wrap gap-3 items-end mt-5">
              <Field label="Material">
                <select style={{ ...inputStyle, width: 260 }} value={kardexMaterial} onChange={(e) => setKardexMaterial(e.target.value)}>
                  {materials.map((m) => <option key={m.id} value={m.id}>{m.nombre}</option>)}
                </select>
              </Field>
            </div>

            {kardexMat && (
              <div className="grid grid-cols-3 gap-4 mt-4">
                <KPICard label="Stock actual" value={`${kardexMat.stock} ${kardexMat.unidad}`} />
                <KPICard label="Costo promedio" value={fmt(kardexMat.costoUnitario)} sub="por unidad" />
                <KPICard label="Valor en inventario" value={fmt(kardexMat.stock * kardexMat.costoUnitario)} accent={C.cyanDark} />
              </div>
            )}

            <div className="flex gap-2 mt-4">
              <button onClick={() => setKardexAction(kardexAction === 'entrada' ? null : 'entrada')} className="flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.cyanDark }}>
                {kardexAction === 'entrada' ? <X size={15} /> : <Plus size={15} />} Registrar entrada (compra)
              </button>
              <button onClick={() => setKardexAction(kardexAction === 'salida' ? null : 'salida')} className="flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.coral }}>
                {kardexAction === 'salida' ? <X size={15} /> : <Plus size={15} />} Registrar salida (consumo)
              </button>
            </div>

            {kardexAction === 'entrada' && <EntradaForm materials={materials} defaultMaterial={kardexMaterial} centros={centros} onSubmit={registerEntrada} />}
            {kardexAction === 'salida' && <SalidaForm materials={materials} defaultMaterial={kardexMaterial} centros={centros} orders={orders.filter((o) => o.estado !== 'cancelada')} onSubmit={registerSalida} />}

            <div className="rounded-xl overflow-hidden mt-5" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
              <table>
                <thead><tr><th>Fecha</th><th>Movimiento</th><th>Referencia</th><th>Centro de costo</th><th>Cantidad</th><th>Costo unitario</th><th>Saldo cantidad</th><th>Saldo valor</th></tr></thead>
                <tbody>
                  {kardexPag.paginaActual.map((k) => (
                    <tr key={k.id}>
                      <td>{k.fecha}</td>
                      <td>
                        <span className="text-xs font-semibold px-2 py-0.5 rounded-full" style={{ backgroundColor: k.tipo === 'entrada' ? '#E4F8F1' : '#FFF4E0', color: k.tipo === 'entrada' ? C.cyanDark : C.amber }}>
                          {k.tipo === 'entrada' ? 'Entrada' : 'Salida'}
                        </span>
                      </td>
                      <td style={{ fontSize: 12.5 }}>{k.referencia}</td>
                      <td style={{ fontSize: 12 }}>{k.centroCostoId ? centroCostoName(k.centroCostoId) : '—'}</td>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{k.tipo === 'entrada' ? '+' : '−'}{k.cantidad}</td>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(k.costoUnitario)}</td>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{k.saldoCantidad}</td>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(k.saldoValor)}</td>
                    </tr>
                  ))}
                  {kardexRows.length === 0 && <tr><td colSpan={8} style={{ color: C.muted, textAlign: 'center', padding: 24 }}>Sin movimientos para este material.</td></tr>}
                </tbody>
              </table>
              <Paginador p={kardexPag} />
            </div>
          </div>
        )}

        {/* FACTURACION */}
        {tab === 'facturacion' && (() => {
          const conDias = completedOrders.map((o) => ({ ...o, dias: Math.floor((new Date() - new Date(o.fecha)) / (1000 * 60 * 60 * 24)) }));
          const pendientes = conDias.filter((o) => o.estadoPago === 'pendiente');
          const vencidas30 = pendientes.filter((o) => o.dias > 30);
          const manualesPend = facturasManuales.filter((f) => f.estadoPago === 'pendiente').reduce((s, f) => s + f.monto, 0);
          return (
          <div>
            <div className="flex items-center justify-between flex-wrap gap-2">
              <div>
                <h1 style={{ fontFamily: 'Space Grotesk, sans-serif', fontSize: 22, fontWeight: 700, color: C.slate }}>{t('Facturación a clientes')}</h1>
                <p style={{ color: C.muted, fontSize: 13, marginTop: 2 }}>El ingreso real del negocio son las <strong>órdenes de trabajo</strong>. Las facturas excepcionales son solo para casos fuera de la operación (venta de sobrante, renta de equipo, etc.).</p>
              </div>
              <div className="flex gap-2">
                <button onClick={() => setShowFacturaManualForm(!showFacturaManualForm)} className="flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-sm font-semibold" style={{ backgroundColor: '#fff', color: C.slate, border: `1px solid ${C.border}` }}>
                  {showFacturaManualForm ? <X size={15} /> : <Plus size={15} />} Factura excepcional
                </button>
                <button onClick={() => descargarCSVInforme('facturas', [['Folio', 'Tipo', 'Referencia', 'Cliente', 'Fecha', 'Días', 'Monto', 'Estado'], ...conDias.map((o) => [o.id, 'Orden', o.referencia || '', clientName(o.clienteId), o.fecha, o.dias, o.precioCliente, o.estadoPago]), ...facturasManuales.map((f) => [f.id, 'Excepcional', f.concepto, clientName(f.clienteId), f.fecha, '', f.monto, f.estadoPago])])}
                  className="flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.navy }}>
                  <Download size={14} /> Exportar CSV
                </button>
              </div>
            </div>

            {showFacturaManualForm && <FacturaManualForm clients={clients.filter((c) => c.activo !== false)} ingresoCategorias={ingresoCategorias} onSubmit={(data) => { addFacturaManual(data); setShowFacturaManualForm(false); }} />}

            <div className="grid grid-cols-3 gap-4 mt-5">
              {[
                ['Por cobrar total', fmt(receivable + manualesPend), C.amber, `${pendientes.length} orden(es) + ${facturasManuales.filter((f) => f.estadoPago === 'pendiente').length} excepcional(es)`],
                ['Vencidas (+30 días)', fmt(vencidas30.reduce((s, o) => s + Number(o.precioCliente), 0)), vencidas30.length > 0 ? C.coral : C.cyanDark, `${vencidas30.length} factura(s) — prioriza su cobro`],
                ['Cobrado histórico', fmt(conDias.filter((o) => o.estadoPago === 'pagada').reduce((s, o) => s + Number(o.precioCliente), 0) + facturasManuales.filter((f) => f.estadoPago === 'pagada').reduce((s, f) => s + f.monto, 0)), C.cyanDark, 'Órdenes + excepcionales'],
              ].map(([label, val, color, desc], i) => (
                <div key={i} className="rounded-xl p-4" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
                  <div style={{ fontSize: 11, fontWeight: 700, color: C.muted, textTransform: 'uppercase' }}>{label}</div>
                  <div style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 20, fontWeight: 700, color, marginTop: 6 }}>{val}</div>
                  <div style={{ fontSize: 11, color: C.muted, marginTop: 4 }}>{desc}</div>
                </div>
              ))}
            </div>

            <div className="rounded-xl overflow-hidden mt-5" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
              <div className="px-4 pt-4 pb-1" style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 600, fontSize: 14, color: C.slate }}>Ingresos por órdenes de trabajo (fuente principal)</div>
              <table>
                <thead><tr><th>Folio</th><th>Cliente</th><th>Fecha</th><th>Antigüedad</th><th>Monto</th><th>Estado de pago</th></tr></thead>
                <tbody>
                  {[...conDias].sort((a, b) => (a.estadoPago === 'pendiente' ? 0 : 1) - (b.estadoPago === 'pendiente' ? 0 : 1) || b.dias - a.dias).map((o) => (
                    <tr key={o.id}>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 12 }}>{o.id}{o.referencia && <div style={{ fontSize: 10, color: C.muted, fontFamily: 'Inter, sans-serif' }}>{o.referencia}</div>}</td>
                      <td>{clientName(o.clienteId)}</td>
                      <td>{o.fecha}</td>
                      <td>
                        {o.estadoPago === 'pagada' ? <span style={{ color: C.muted, fontSize: 12 }}>—</span> : (
                          <span className="text-xs font-semibold px-2 py-0.5 rounded-full" style={{ backgroundColor: o.dias > 60 ? '#FDE8E8' : o.dias > 30 ? '#FDEBD0' : '#E4F8F1', color: o.dias > 60 ? C.coral : o.dias > 30 ? '#8A5A00' : C.cyanDark }}>
                            {o.dias} días
                          </span>
                        )}
                      </td>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(o.precioCliente)}</td>
                      <td><PayPill ok={o.estadoPago === 'pagada'} textOk="Pagada" textPending="Marcar pagada" onClick={() => togglePagoCliente(o.id)} /></td>
                    </tr>
                  ))}
                  {conDias.length === 0 && <tr><td colSpan={6} style={{ color: C.muted, textAlign: 'center', padding: 24 }}>Aún no hay órdenes con ingreso reconocido.</td></tr>}
                </tbody>
                {conDias.length > 0 && (
                  <tfoot><tr><td colSpan={4} style={{ fontWeight: 700 }}>Total ({conDias.length})</td><td style={{ fontFamily: 'JetBrains Mono, monospace', fontWeight: 700 }}>{fmt(conDias.reduce((s, o) => s + Number(o.precioCliente), 0))}</td><td></td></tr></tfoot>
                )}
              </table>
            </div>

            <div className="rounded-xl overflow-hidden mt-5" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
              <div className="px-4 pt-4 pb-1" style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 600, fontSize: 14, color: C.slate }}>Facturas excepcionales (casos fuera de la operación)</div>
              <table>
                <thead><tr><th>Folio</th><th>Cliente</th><th>Concepto</th><th>Categoría de ingreso</th><th>Fecha</th><th>Monto</th><th>Estado</th></tr></thead>
                <tbody>
                  {facturasManuales.map((f) => (
                    <tr key={f.id}>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 12 }}>{f.id}</td>
                      <td>{clientName(f.clienteId)}</td>
                      <td style={{ fontSize: 12.5 }}>{f.concepto}</td>
                      <td style={{ fontSize: 12 }}>{f.categoria} <span style={{ color: C.muted }}>({f.cuentaIngreso})</span></td>
                      <td>{f.fecha}</td>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(f.monto)}</td>
                      <td><PayPill ok={f.estadoPago === 'pagada'} textOk="Cobrada" textPending="Marcar cobrada" onClick={() => toggleFacturaManualPago(f.id)} /></td>
                    </tr>
                  ))}
                  {facturasManuales.length === 0 && <tr><td colSpan={7} style={{ color: C.muted, textAlign: 'center', padding: 20 }}>Sin facturas excepcionales — así debe ser la mayoría del tiempo.</td></tr>}
                </tbody>
                {facturasManuales.length > 0 && (
                  <tfoot><tr><td colSpan={5} style={{ fontWeight: 700 }}>Total ({facturasManuales.length})</td><td style={{ fontFamily: 'JetBrains Mono, monospace', fontWeight: 700 }}>{fmt(facturasManuales.reduce((s, f) => s + f.monto, 0))}</td><td></td></tr></tfoot>
                )}
              </table>
            </div>
          </div>
          );
        })()}

        {/* PAGOS */}
        {/* BONOS POR PRODUCTIVIDAD */}
        {tab === 'bonos' && (() => {
          const desde = bonoMes + '-01';
          const hasta = bonoMes + '-31';
          const filas = techs.map((t) => {
            const completadas = orders.filter((o) => o.tecnicoId === t.id && o.estado === 'completada' && o.fecha >= desde && o.fecha <= hasta).length;
            const periodoBono = `Bono ${bonoMes}`;
            const yaOtorgado = nomina.some((n) => n.empleadoId === t.id && n.periodo === periodoBono);
            return { ...t, completadas, elegible: completadas >= Number(bonoUmbral), yaOtorgado, periodoBono };
          });
          const otorgarBono = (t) => {
            if (t.yaOtorgado) return;
            addNomina({ empleadoId: t.id, periodo: t.periodoBono, fechaPago: new Date().toISOString().slice(0, 10), monto: Number(bonoMonto) || 0 });
            alert(`Bono de ${fmt(Number(bonoMonto))} otorgado a ${t.nombre} por ${t.completadas} instalaciones en ${bonoMes}.`);
          };
          return (
            <div>
              <h1 style={{ fontFamily: 'Space Grotesk, sans-serif', fontSize: 22, fontWeight: 700, color: C.slate }}>{t('Bonos por productividad')}</h1>
              <p style={{ color: C.muted, fontSize: 13, marginTop: 2 }}>Al alcanzar la meta de instalaciones completadas en el mes, el técnico recibe un bono. Se registra en la nómina y se contabiliza automáticamente.</p>

              <div className="flex flex-wrap gap-3 items-end mt-5">
                <Field label="Mes"><input type="month" style={inputStyle} value={bonoMes} onChange={(e) => setBonoMes(e.target.value)} /></Field>
                <Field label="Meta de instalaciones"><input type="number" style={{ ...inputStyle, width: 130 }} value={bonoUmbral} onChange={(e) => setBonoUmbral(e.target.value)} /></Field>
                <Field label="Monto del bono ($)"><input type="number" style={{ ...inputStyle, width: 130 }} value={bonoMonto} onChange={(e) => setBonoMonto(e.target.value)} /></Field>
              </div>

              <div className="rounded-xl overflow-hidden mt-5" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
                <table>
                  <thead><tr><th>Técnico</th><th>Instalaciones completadas en {bonoMes}</th><th>Avance a la meta</th><th>Estado</th><th></th></tr></thead>
                  <tbody>
                    {filas.map((t) => (
                      <tr key={t.id}>
                        <td style={{ fontWeight: 600 }}>{t.nombre}</td>
                        <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{t.completadas} / {bonoUmbral}</td>
                        <td>
                          <div style={{ width: 140, height: 9, backgroundColor: '#EEF1F4', borderRadius: 5, overflow: 'hidden' }}>
                            <div style={{ width: `${Math.min(100, (t.completadas / Math.max(1, Number(bonoUmbral))) * 100)}%`, height: '100%', backgroundColor: t.elegible ? C.cyanDark : C.amber }} />
                          </div>
                        </td>
                        <td>
                          {t.yaOtorgado ? <span className="text-xs font-semibold px-2 py-0.5 rounded-full" style={{ backgroundColor: '#E4F8F1', color: C.cyanDark }}>✓ Bono otorgado</span>
                            : t.elegible ? <span className="text-xs font-semibold px-2 py-0.5 rounded-full" style={{ backgroundColor: '#FDEBD0', color: '#8A5A00' }}>Elegible</span>
                            : <span style={{ color: C.muted, fontSize: 12 }}>Faltan {Math.max(0, Number(bonoUmbral) - t.completadas)}</span>}
                        </td>
                        <td>
                          {t.elegible && !t.yaOtorgado && (
                            <button onClick={() => otorgarBono(t)} className="px-3 py-1.5 rounded-lg text-white text-xs font-semibold" style={{ backgroundColor: C.cyanDark }}>Otorgar bono</button>
                          )}
                        </td>
                      </tr>
                    ))}
                    {filas.length === 0 && <tr><td colSpan={5} style={{ color: C.muted, textAlign: 'center', padding: 20 }}>Sin técnicos registrados.</td></tr>}
                  </tbody>
                </table>
              </div>
              <div style={{ fontSize: 11.5, color: C.muted, marginTop: 10 }}>El bono se registra como nómina del técnico (periodo "Bono {bonoMes}") y genera su asiento contable. No se puede otorgar dos veces el mismo mes.</div>
            </div>
          );
        })()}

        {/* CUENTAS POR PAGAR */}
        {tab === 'cxp' && (
          <div>
            <div className="flex items-center justify-between">
              <div>
                <h1 style={{ fontFamily: 'Space Grotesk, sans-serif', fontSize: 22, fontWeight: 700, color: C.slate }}>{t('Cuentas por pagar')}</h1>
                <p style={{ color: C.muted, fontSize: 13, marginTop: 2 }}>Facturas y compras a proveedores. Pendiente: <strong style={{ color: C.amber }}>{fmt(cxpPendiente)}</strong></p>
              </div>
              <div className="flex gap-2">
                <button onClick={() => setShowProvForm(!showProvForm)} className="flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-sm font-semibold" style={{ backgroundColor: '#fff', color: C.slate, border: `1px solid ${C.border}` }}>
                  {showProvForm ? <X size={15} /> : <Plus size={15} />} Proveedor
                </button>
                <button onClick={() => setShowCxpForm(!showCxpForm)} className="flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.navy }}>
                  {showCxpForm ? <X size={15} /> : <Plus size={15} />} Cuenta por pagar
                </button>
              </div>
            </div>

            {showProvForm && (
              <SimpleForm fields={[{ key: 'nombre', label: 'Nombre del proveedor' }, { key: 'contacto', label: 'Contacto' }, { key: 'telefono', label: 'Teléfono' }]}
                onSubmit={(data) => { addProveedor(data); setShowProvForm(false); }} />
            )}
            {showCxpForm && <CxpForm proveedores={proveedores.filter((p) => p.activo !== false)} gastoCategorias={gastoCategorias} materials={materials} onSubmit={(data) => { addCxp(data); setShowCxpForm(false); }} />}

            <div className="rounded-xl overflow-hidden mt-5" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
              <table>
                <thead><tr><th>Proveedor</th><th>Concepto</th><th>Tipo</th><th>Fecha</th><th>Vencimiento</th><th>Monto</th><th>Estado</th></tr></thead>
                <tbody>
                  {cxpPag.paginaActual.map((c) => (
                    <tr key={c.id}>
                      <td>{proveedorName(c.proveedorId)}</td>
                      <td>{c.concepto}</td>
                      <td><span className="text-xs font-semibold px-2 py-0.5 rounded-full" style={{ backgroundColor: c.tipo === 'inventario' ? '#E4F8F1' : '#EEF1F4', color: c.tipo === 'inventario' ? C.cyanDark : C.slate }}>{c.tipo === 'inventario' ? 'Inventario' : 'Gasto'}</span></td>
                      <td>{c.fecha}</td>
                      <td>{c.vencimiento}</td>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(c.monto)}</td>
                      <td><PayPill ok={c.estado === 'pagada'} textOk="Pagada" textPending="Marcar pagada" onClick={() => toggleCxpPago(c.id)} /></td>
                    </tr>
                  ))}
                  {cxp.length === 0 && <tr><td colSpan={7} style={{ color: C.muted, textAlign: 'center', padding: 24 }}>Sin cuentas por pagar registradas.</td></tr>}
                </tbody>
                {cxp.length > 0 && (
                  <tfoot><tr><td colSpan={5} style={{ fontWeight: 700 }}>Total ({cxp.length})</td><td style={{ fontFamily: 'JetBrains Mono, monospace', fontWeight: 700 }}>{fmt(cxp.reduce((s, c) => s + c.monto, 0))}</td><td></td></tr></tfoot>
                )}
              </table>
              <Paginador p={cxpPag} />
            </div>
          </div>
        )}

        {/* PROVEEDORES */}
        {tab === 'proveedores' && (
          <div>
            <div className="flex items-center justify-between">
              <h1 style={{ fontFamily: 'Space Grotesk, sans-serif', fontSize: 22, fontWeight: 700, color: C.slate }}>Proveedores</h1>
              <button onClick={() => setShowProvForm(!showProvForm)} className="flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.navy }}>
                {showProvForm ? <X size={15} /> : <Plus size={15} />} {showProvForm ? 'Cerrar' : 'Nuevo proveedor'}
              </button>
            </div>
            {showProvForm && (
              <SimpleForm fields={[{ key: 'nombre', label: 'Nombre del proveedor' }, { key: 'contacto', label: 'Contacto' }, { key: 'telefono', label: 'Teléfono' }]}
                onSubmit={(data) => { addProveedor(data); setShowProvForm(false); }} />
            )}
            <div className="grid grid-cols-2 gap-3 mt-5">
              {proveedores.map((p) => {
                const totalCompras = cxp.filter((c) => c.proveedorId === p.id).reduce((s, c) => s + Number(c.monto), 0);
                const pendiente = cxp.filter((c) => c.proveedorId === p.id && c.estado === 'pendiente').reduce((s, c) => s + Number(c.monto), 0);
                return (
                  <div key={p.id} className="rounded-xl p-4" style={{ backgroundColor: C.card, border: `1px solid ${C.border}`, opacity: p.activo === false ? 0.55 : 1 }}>
                    <div className="flex items-center justify-between">
                      <input value={p.nombre} onChange={(ev) => updateProveedor(p.id, 'nombre', ev.target.value)} style={{ ...inputStyle, fontWeight: 600, fontSize: 13.5, padding: '3px 6px', width: 200 }} />
                      <div className="flex items-center gap-2">
                        {p.activo === false && <span className="text-xs font-bold px-2 py-0.5 rounded-full" style={{ backgroundColor: '#FDE8E8', color: C.coral }}>INACTIVO</span>}
                        <button onClick={() => toggleProveedorActivo(p.id)} className="text-xs font-semibold px-2 py-1 rounded-lg" style={{ backgroundColor: p.activo === false ? '#E4F8F1' : '#FDE8E8', color: p.activo === false ? C.cyanDark : C.coral }}>
                          {p.activo === false ? 'Reactivar' : 'Dar de baja'}
                        </button>
                      </div>
                    </div>
                    <div className="flex items-center gap-1.5 mt-2" style={{ color: C.muted, fontSize: 12.5 }}><span>Contacto:</span><input value={p.contacto || ''} onChange={(ev) => updateProveedor(p.id, 'contacto', ev.target.value)} style={{ ...inputStyle, fontSize: 11.5, padding: '3px 6px', width: 180 }} /></div>
                    <div className="flex items-center gap-1.5 mt-1" style={{ color: C.muted, fontSize: 12.5 }}><Phone size={12} /><input value={p.telefono || ''} onChange={(ev) => updateProveedor(p.id, 'telefono', ev.target.value)} style={{ ...inputStyle, fontSize: 11.5, padding: '3px 6px', width: 140 }} /></div>
                    <div className="flex gap-3 mt-2.5 text-xs" style={{ fontFamily: 'JetBrains Mono, monospace', fontWeight: 600 }}>
                      <span style={{ color: C.slate }}>Comprado: {fmt(totalCompras)}</span>
                      {pendiente > 0 && <span style={{ color: C.amber }}>Pendiente: {fmt(pendiente)}</span>}
                    </div>
                  </div>
                );
              })}
              {proveedores.length === 0 && <div style={{ color: C.muted, fontSize: 13, gridColumn: '1 / -1', textAlign: 'center', padding: 24 }}>Sin proveedores registrados.</div>}
            </div>
          </div>
        )}

        {/* ACTIVOS FIJOS */}
        {tab === 'activos' && (
          <div>
            <div className="flex items-center justify-between">
              <div>
                <h1 style={{ fontFamily: 'Space Grotesk, sans-serif', fontSize: 22, fontWeight: 700, color: C.slate }}>{t('Activos fijos')}</h1>
                <p style={{ color: C.muted, fontSize: 13, marginTop: 2 }}>Vehículos, equipo y herramienta con depreciación en línea recta. Valor en libros estimado: <strong style={{ color: C.cyanDark }}>{fmt(valorActivosLibros)}</strong></p>
              </div>
              <button onClick={() => setShowActivoForm(!showActivoForm)} className="flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.navy }}>
                {showActivoForm ? <X size={15} /> : <Plus size={15} />} {showActivoForm ? 'Cerrar' : 'Nuevo activo'}
              </button>
            </div>

            {showActivoForm && <ActivoForm techs={techs} centros={centros} onSubmit={(data) => { addActivo(data); setShowActivoForm(false); }} />}

            <div className="rounded-xl overflow-hidden mt-5" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
              <table>
                <thead><tr><th>Activo</th><th>Categoría</th><th>Centro de costo</th><th>Responsable</th><th>Fecha compra</th><th>Costo</th><th>Dep. mensual</th><th>Dep. contabilizada</th><th>Valor en libros contable</th></tr></thead>
                <tbody>
                  {activos.map((a) => {
                    const d = depreciacionInfo(a);
                    const contab = depContabilizada(a.id);
                    return (
                      <tr key={a.id}>
                        <td>{a.nombre}</td>
                        <td>{a.categoria}</td>
                        <td style={{ fontSize: 12 }}>{centroCostoName(a.centroCostoId)}</td>
                        <td>{techName(a.responsable)}</td>
                        <td>{a.fechaCompra}</td>
                        <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(a.costo)}</td>
                        <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(d.depAnual / 12)}</td>
                        <td style={{ fontFamily: 'JetBrains Mono, monospace', color: C.amber }}>{fmt(contab)}</td>
                        <td style={{ fontFamily: 'JetBrains Mono, monospace', color: C.cyanDark }}>{fmt(a.costo - contab)}</td>
                      </tr>
                    );
                  })}
                  {activos.length === 0 && <tr><td colSpan={9} style={{ color: C.muted, textAlign: 'center', padding: 24 }}>Sin activos registrados.</td></tr>}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* CONTROL DE DEPRECIACIÓN — página separada, dentro del grupo Activos fijos */}
        {tab === 'depreciacion' && (
          <div>
            <h1 style={{ fontFamily: 'Space Grotesk, sans-serif', fontSize: 22, fontWeight: 700, color: C.slate }}>Control de depreciación</h1>
            <p style={{ color: C.muted, fontSize: 13, marginTop: 2 }}>Registra la depreciación mensual de todos los activos fijos y consulta su historial.</p>

            <div className="rounded-xl p-4 mt-5 flex flex-wrap gap-3 items-end" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
              <Field label="Mes a registrar"><input type="month" style={inputStyle} value={depMes} onChange={(e) => setDepMes(e.target.value)} /></Field>
              <button onClick={() => registrarDepreciacionMes(depMes)} className="px-4 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.cyanDark }}>Registrar depreciación del mes</button>
              <div style={{ fontSize: 11.5, color: C.muted, width: '100%' }}>Genera un asiento por cada activo en su centro de costo: los operativos van a Costo de depreciación (6310) y los del centro Administrativo a Gasto de depreciación (5310). No duplica un mes ya registrado.</div>
            </div>

            <div className="rounded-xl overflow-hidden mt-5" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
              <div className="px-4 pt-4 pb-1" style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 600, fontSize: 14, color: C.slate }}>Historial de depreciación registrada</div>
              <table>
                <thead><tr><th>Periodo</th><th>Activo</th><th>Centro de costo</th><th>Monto</th></tr></thead>
                <tbody>
                  {[...depreciaciones].sort((a, b) => b.periodo.localeCompare(a.periodo)).map((d) => {
                    const a = activos.find((x) => x.id === d.activoId);
                    return (
                      <tr key={d.id}>
                        <td>{d.periodo}</td>
                        <td>{a?.nombre || '—'}</td>
                        <td style={{ fontSize: 12 }}>{a ? centroCostoName(a.centroCostoId) : '—'}</td>
                        <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(d.monto)}</td>
                      </tr>
                    );
                  })}
                  {depreciaciones.length === 0 && <tr><td colSpan={4} style={{ color: C.muted, textAlign: 'center', padding: 20 }}>Aún no se ha registrado ninguna depreciación mensual.</td></tr>}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* FLOTA — página separada, dentro del grupo Activos fijos: combustible, mantenimiento, seguro/revisión, multas */}
        {tab === 'flota' && !activos.some((a) => a.categoria === 'Vehículo') && (
          <div>
            <h1 style={{ fontFamily: 'Space Grotesk, sans-serif', fontSize: 22, fontWeight: 700, color: C.slate }}>Bitácora de flota</h1>
            <div className="rounded-xl p-6 mt-5 text-center" style={{ backgroundColor: C.card, border: `1px solid ${C.border}`, color: C.muted, fontSize: 13 }}>
              Aún no tienes ningún activo de categoría "Vehículo". Da de alta uno en <button onClick={() => setTab('activos')} style={{ color: C.cyanDark, fontWeight: 600, textDecoration: 'underline' }}>Activos fijos</button> para empezar a llevar su bitácora aquí.
            </div>
          </div>
        )}
        {tab === 'flota' && activos.some((a) => a.categoria === 'Vehículo') && (
          <div>
            <h1 style={{ fontFamily: 'Space Grotesk, sans-serif', fontSize: 22, fontWeight: 700, color: C.slate }}>Bitácora de flota</h1>
            <p style={{ color: C.muted, fontSize: 13, marginTop: 2 }}>Vencimientos de seguro y revisión, y bitácora de combustible, mantenimiento y multas por unidad.</p>

            <div className="rounded-xl overflow-hidden mt-4" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
              <table>
                <thead><tr><th>Vehículo</th><th>Placa</th><th>Km</th><th>Vence seguro</th><th>Vence revisión</th><th>Gasto del mes</th><th></th></tr></thead>
                <tbody>
                  {activos.filter((a) => a.categoria === 'Vehículo').map((a) => {
                    const hoy = new Date().toISOString().slice(0, 10);
                    const mesActual = hoy.slice(0, 7);
                    const gastoMes = flotaLog.filter((f) => f.activoId === a.id && f.fecha?.startsWith(mesActual)).reduce((s, f) => s + f.monto, 0);
                    const vencePronto = (fecha) => fecha && fecha <= new Date(Date.now() + 30 * 86400000).toISOString().slice(0, 10);
                    const vencido = (fecha) => fecha && fecha < hoy;
                    const badgeVenc = (fecha) => !fecha ? <span style={{ color: C.muted, fontSize: 11.5 }}>—</span> : (
                      <span className="text-xs font-semibold px-2 py-0.5 rounded-full" style={{ backgroundColor: vencido(fecha) ? '#FDE8E8' : vencePronto(fecha) ? '#FFF4E0' : '#E4F8F1', color: vencido(fecha) ? C.coral : vencePronto(fecha) ? '#8A5A00' : C.cyanDark }}>
                        {fecha} {vencido(fecha) ? '⚠ vencido' : vencePronto(fecha) ? '⚠ pronto' : ''}
                      </span>
                    );
                    return (
                      <tr key={a.id}>
                        <td>{a.nombre}</td>
                        <td style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 12 }}>{a.placa || '—'}</td>
                        <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{a.km ?? '—'}</td>
                        <td>{badgeVenc(a.seguroVencimiento)}</td>
                        <td>{badgeVenc(a.revisionVencimiento)}</td>
                        <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(gastoMes)}</td>
                        <td>
                          <button onClick={() => { if (showFlotaForm && flotaActivoSel === a.id) { setShowFlotaForm(null); } else { setFlotaActivoSel(a.id); setShowFlotaForm('combustible'); } }} className="text-xs font-semibold px-2 py-1 rounded-lg" style={{ backgroundColor: '#EEF1F4', color: C.slate }}>
                            + Registro
                          </button>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>

            {showFlotaForm && flotaActivoSel && (
              <FlotaForm activoNombre={activos.find((a) => a.id === flotaActivoSel)?.nombre}
                onSubmit={(tipo, data) => { addFlotaRegistro(tipo, { ...data, activoId: flotaActivoSel }); setShowFlotaForm(null); }}
                onClose={() => setShowFlotaForm(null)} />
            )}

            <div className="rounded-xl overflow-hidden mt-5" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
              <div className="px-4 pt-4 pb-1" style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 600, fontSize: 14, color: C.slate }}>Bitácora de flota</div>
              <table>
                <thead><tr><th>Fecha</th><th>Vehículo</th><th>Tipo</th><th>Detalle</th><th>Monto</th></tr></thead>
                <tbody>
                  {[...flotaLog].reverse().map((f) => (
                    <tr key={f.id}>
                      <td>{f.fecha}</td>
                      <td>{activos.find((a) => a.id === f.activoId)?.nombre || '—'}</td>
                      <td>
                        <span className="text-xs font-semibold px-2 py-0.5 rounded-full" style={{
                          backgroundColor: f.tipo === 'combustible' ? '#E4F8F1' : f.tipo === 'mantenimiento' ? '#FFF4E0' : '#FDE8E8',
                          color: f.tipo === 'combustible' ? C.cyanDark : f.tipo === 'mantenimiento' ? '#8A5A00' : C.coral,
                        }}>
                          {f.tipo === 'combustible' ? 'Combustible' : f.tipo === 'mantenimiento' ? `Mantenimiento (${f.subtipo})` : 'Multa'}
                        </span>
                      </td>
                      <td style={{ fontSize: 12 }}>{f.detalle || '—'}</td>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(f.monto)}</td>
                    </tr>
                  ))}
                  {flotaLog.length === 0 && <tr><td colSpan={5} style={{ color: C.muted, textAlign: 'center', padding: 20 }}>Sin registros de flota todavía.</td></tr>}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* RECURSOS HUMANOS */}
        {tab === 'rh' && (
          <div>
            <div className="flex items-center justify-between">
              <div>
                <h1 style={{ fontFamily: 'Space Grotesk, sans-serif', fontSize: 22, fontWeight: 700, color: C.slate }}>{t('Recursos humanos')}</h1>
                <p style={{ color: C.muted, fontSize: 13, marginTop: 2 }}>Todo el personal: administrativos y técnicos. Nómina mensual: <strong style={{ color: C.slate }}>{fmt(nominaMensual)}</strong> · Pendiente por pagar: <strong style={{ color: C.amber }}>{fmt(nominaPendiente)}</strong></p>
              </div>
              <div className="flex gap-2 items-end flex-wrap">
                <button onClick={() => setShowEmpleadoForm(!showEmpleadoForm)} className="flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-sm font-semibold" style={{ backgroundColor: '#fff', color: C.slate, border: `1px solid ${C.border}` }}>
                  {showEmpleadoForm ? <X size={15} /> : <Plus size={15} />} Empleado
                </button>
                <button onClick={() => setShowNominaForm(!showNominaForm)} className="flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-sm font-semibold" style={{ backgroundColor: '#fff', color: C.slate, border: `1px solid ${C.border}` }}>
                  {showNominaForm ? <X size={15} /> : <Plus size={15} />} Nómina individual
                </button>
                <button onClick={() => setShowAnticipoForm(!showAnticipoForm)} className="flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-sm font-semibold" style={{ backgroundColor: '#fff', color: C.slate, border: `1px solid ${C.border}` }}>
                  {showAnticipoForm ? <X size={15} /> : <Plus size={15} />} Anticipo
                </button>
                <button onClick={() => setShowPrestamoForm(!showPrestamoForm)} className="flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-sm font-semibold" style={{ backgroundColor: '#fff', color: C.slate, border: `1px solid ${C.border}` }}>
                  {showPrestamoForm ? <X size={15} /> : <Plus size={15} />} Préstamo
                </button>
                <Field label="Mes"><input type="month" style={{ ...inputStyle, width: 140 }} value={nominaMes} onChange={(e) => setNominaMes(e.target.value)} /></Field>
                <Field label="Ciclo">
                  <select style={{ ...inputStyle, width: 130 }} value={nominaCiclo} onChange={(e) => setNominaCiclo(e.target.value)}>
                    <option value="mensual">Mensual</option>
                    <option value="quincena1">1ra quincena</option>
                    <option value="quincena2">2da quincena</option>
                  </select>
                </Field>
                <button onClick={() => registrarNominaMes(nominaMes, nominaCiclo)} className="flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.cyanDark }}>
                  Registrar nómina
                </button>
              </div>
            </div>

            {showAnticipoForm && <AnticipoForm empleados={empleados.filter((e) => e.activo !== false)} onSubmit={(data) => { addAnticipo(data); setShowAnticipoForm(false); }} />}
            {showPrestamoForm && <PrestamoForm empleados={empleados.filter((e) => e.activo !== false)} onSubmit={(data) => { addPrestamo(data); setShowPrestamoForm(false); }} />}

            {showEmpleadoForm && (
              <SimpleForm fields={[
                { key: 'nombre', label: 'Nombre' },
                { key: 'rol', label: 'Rol', type: 'select', options: ['Administrativo', 'Técnico'] },
                { key: 'puesto', label: 'Puesto / especialidad' },
                { key: 'telefono', label: 'Teléfono' },
                { key: 'fechaIngreso', label: 'Fecha de ingreso', type: 'date' },
                { key: 'salarioMensual', label: 'Salario mensual ($)', type: 'number' },
              ]}
                onSubmit={(data) => { addEmpleado(data); setShowEmpleadoForm(false); }} />
            )}
            {showNominaForm && <NominaForm empleados={empleados} onSubmit={(data) => { addNomina(data); setShowNominaForm(false); }} />}

            <div className="grid grid-cols-2 gap-3 mt-5">
              {empleados.map((e) => (
                <div key={e.id} className="rounded-xl p-4" style={{ backgroundColor: C.card, border: `1px solid ${C.border}`, opacity: e.activo === false ? 0.55 : 1 }}>
                  <div className="flex items-center justify-between">
                    <div className="flex items-center gap-2">
                      <span style={{ fontWeight: 600, fontSize: 14, color: C.slate }}>{e.nombre}</span>
                      <span className="text-xs font-semibold px-2 py-0.5 rounded-full" style={{ backgroundColor: e.rol === 'Técnico' ? '#E4F8F1' : '#EEF1F4', color: e.rol === 'Técnico' ? C.cyanDark : C.slate }}>{e.rol || 'Administrativo'}</span>
                      {e.activo === false && <span className="text-xs font-bold px-2 py-0.5 rounded-full" style={{ backgroundColor: '#FDE8E8', color: C.coral }}>INACTIVO</span>}
                    </div>
                    <button onClick={() => toggleEmpleadoActivo(e.id)} className="text-xs font-semibold px-2 py-1 rounded-lg" style={{ backgroundColor: e.activo === false ? '#E4F8F1' : '#FDE8E8', color: e.activo === false ? C.cyanDark : C.coral }}>
                      {e.activo === false ? 'Reactivar' : 'Dar de baja'}
                    </button>
                  </div>
                  <div className="flex gap-2 mt-2 flex-wrap items-center" style={{ fontSize: 12 }}>
                    <input value={e.puesto || ''} onChange={(ev) => updateEmpleado(e.id, 'puesto', ev.target.value)} style={{ ...inputStyle, fontSize: 11.5, padding: '3px 6px', width: 170 }} placeholder="Puesto" />
                    <input value={e.telefono || ''} onChange={(ev) => updateEmpleado(e.id, 'telefono', ev.target.value)} style={{ ...inputStyle, fontSize: 11.5, padding: '3px 6px', width: 120 }} placeholder="Teléfono" />
                  </div>
                  <div className="flex gap-3 mt-2 items-center" style={{ fontSize: 11.5, color: C.muted }}>
                    <span>Salario $<input type="number" value={e.salarioMensual || 0} onChange={(ev) => updateEmpleado(e.id, 'salarioMensual', ev.target.value)} style={{ ...inputStyle, fontSize: 11.5, padding: '3px 5px', width: 85, fontFamily: 'JetBrains Mono, monospace' }} />/mes</span>
                    <span>Ingreso: {e.fechaIngreso}</span>
                  </div>
                  <div className="flex gap-2 mt-2 items-center" style={{ fontSize: 11.5, color: C.muted }}>
                    <span>Reporta a:</span>
                    <select value={e.reportaA || ''} onChange={(ev) => updateEmpleado(e.id, 'reportaA', ev.target.value || null)} style={{ ...inputStyle, fontSize: 11.5, padding: '3px 5px', width: 160 }}>
                      <option value="">— Nadie (nivel superior) —</option>
                      {empleados.filter((o) => o.id !== e.id).map((o) => <option key={o.id} value={o.id}>{o.nombre}</option>)}
                    </select>
                  </div>
                </div>
              ))}
            </div>

            {/* SQUADRAS — equipos de técnicos, dentro de RH */}
            <div className="mt-8 pt-6" style={{ borderTop: `1px solid ${C.border}` }}>
              <div className="flex items-center justify-between">
                <div>
                  <h2 style={{ fontFamily: 'Space Grotesk, sans-serif', fontSize: 17, fontWeight: 700, color: C.slate }}>Squadras (equipos de campo)</h2>
                  <p style={{ color: C.muted, fontSize: 12.5, marginTop: 2 }}>Agrupa técnicos en cuadrillas para asignar la agenda por equipo, no solo por persona.</p>
                </div>
                <button onClick={() => setShowSquadraForm(!showSquadraForm)} className="flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.navy }}>
                  {showSquadraForm ? <X size={15} /> : <Plus size={15} />} {showSquadraForm ? 'Cerrar' : 'Nueva squadra'}
                </button>
              </div>

              {showSquadraForm && <SquadraForm techs={techs} onSubmit={(data) => { addSquadra(data); setShowSquadraForm(false); }} />}

              <div className="grid grid-cols-2 gap-3 mt-4">
                {squadras.map((s) => (
                  <div key={s.id} className="rounded-xl p-4" style={{ backgroundColor: C.card, border: `1px solid ${C.border}`, opacity: s.activa === false ? 0.55 : 1 }}>
                    <div className="flex items-center justify-between">
                      <span style={{ fontWeight: 600, fontSize: 14, color: C.slate }}>{s.nombre}</span>
                      <button onClick={() => toggleSquadraActiva(s.id)} className="text-xs font-semibold px-2 py-1 rounded-lg" style={{ backgroundColor: s.activa === false ? '#E4F8F1' : '#FDE8E8', color: s.activa === false ? C.cyanDark : C.coral }}>
                        {s.activa === false ? 'Reactivar' : 'Desactivar'}
                      </button>
                    </div>
                    <div style={{ fontSize: 12, color: C.muted, marginTop: 6 }}>
                      {s.miembros.map((mId) => empleadoName(mId)).join(', ') || 'Sin integrantes'}
                    </div>
                  </div>
                ))}
                {squadras.length === 0 && <div style={{ color: C.muted, fontSize: 13, padding: 12 }}>Sin squadras creadas todavía.</div>}
              </div>
            </div>

            <div className="rounded-xl overflow-hidden mt-6" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
              <div className="px-4 pt-4 pb-1" style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 600, fontSize: 14, color: C.slate }}>Historial de nómina</div>
              <table>
                <thead><tr><th>Empleado</th><th>Periodo</th><th>Fecha de pago</th><th>Bruto</th><th>Descuentos</th><th>Neto a pagar</th><th>Estado</th></tr></thead>
                <tbody>
                  {nominaPag.paginaActual.map((n) => (
                    <tr key={n.id}>
                      <td>{empleadoName(n.empleadoId)}</td>
                      <td>{n.periodo}</td>
                      <td>{n.fechaPago}</td>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(n.montoBruto ?? n.monto)}</td>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace', color: (n.descuentoAnticipo || n.descuentoPrestamo) ? C.coral : C.muted }}>
                        {fmt((n.descuentoAnticipo || 0) + (n.descuentoPrestamo || 0))}
                      </td>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace', fontWeight: 600 }}>{fmt(n.monto)}</td>
                      <td><PayPill ok={n.estado === 'pagado'} textOk="Pagado" textPending="Marcar pagado" onClick={() => toggleNominaPago(n.id)} /></td>
                    </tr>
                  ))}
                  {nomina.length === 0 && <tr><td colSpan={7} style={{ color: C.muted, textAlign: 'center', padding: 20 }}>Sin nómina registrada todavía.</td></tr>}
                </tbody>
                {nomina.length > 0 && (
                  <tfoot>
                    <tr>
                      <td colSpan={3} style={{ fontWeight: 700 }}>Total ({nomina.length})</td>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace', fontWeight: 700 }}>{fmt(nomina.reduce((s, n) => s + (n.montoBruto ?? n.monto), 0))}</td>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace', fontWeight: 700 }}>{fmt(nomina.reduce((s, n) => s + (n.descuentoAnticipo || 0) + (n.descuentoPrestamo || 0), 0))}</td>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace', fontWeight: 700 }}>{fmt(nomina.reduce((s, n) => s + n.monto, 0))}</td>
                      <td></td>
                    </tr>
                  </tfoot>
                )}
              </table>
              <Paginador p={nominaPag} />
            </div>

            {/* ANTICIPOS DE SUELDO */}
            <div className="rounded-xl overflow-hidden mt-6" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
              <div className="px-4 pt-4 pb-1" style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 600, fontSize: 14, color: C.slate }}>Anticipos de sueldo</div>
              <p style={{ fontSize: 11.5, color: C.muted, padding: '0 16px 8px' }}>Se descuentan automáticamente y por completo en la siguiente nómina del empleado.</p>
              <table>
                <thead><tr><th>Empleado</th><th>Fecha</th><th>Motivo</th><th>Monto</th><th>Estado</th></tr></thead>
                <tbody>
                  {anticipos.map((a) => (
                    <tr key={a.id}>
                      <td>{empleadoName(a.empleadoId)}</td>
                      <td>{a.fecha}</td>
                      <td style={{ fontSize: 12.5 }}>{a.motivo || '—'}</td>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(a.monto)}</td>
                      <td>
                        <span className="text-xs font-semibold px-2 py-0.5 rounded-full" style={{ backgroundColor: a.estado === 'pendiente' ? '#FFF4E0' : '#E4F8F1', color: a.estado === 'pendiente' ? '#8A5A00' : C.cyanDark }}>
                          {a.estado === 'pendiente' ? 'Pendiente de descontar' : 'Descontado'}
                        </span>
                      </td>
                    </tr>
                  ))}
                  {anticipos.length === 0 && <tr><td colSpan={5} style={{ color: C.muted, textAlign: 'center', padding: 18 }}>Sin anticipos registrados.</td></tr>}
                </tbody>
                {anticipos.length > 0 && (
                  <tfoot><tr><td colSpan={3} style={{ fontWeight: 700 }}>Pendiente por descontar</td><td colSpan={2} style={{ fontFamily: 'JetBrains Mono, monospace', fontWeight: 700, color: C.amber }}>{fmt(anticipos.filter((a) => a.estado === 'pendiente').reduce((s, a) => s + a.monto, 0))}</td></tr></tfoot>
                )}
              </table>
            </div>

            {/* PRESTAMOS */}
            <div className="rounded-xl overflow-hidden mt-6" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
              <div className="px-4 pt-4 pb-1" style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 600, fontSize: 14, color: C.slate }}>Préstamos a empleados</div>
              <p style={{ fontSize: 11.5, color: C.muted, padding: '0 16px 8px' }}>Se descuenta una cuota automáticamente en cada nómina del empleado hasta liquidarse.</p>
              <table>
                <thead><tr><th>Empleado</th><th>Fecha</th><th>Motivo</th><th>Total</th><th>Cuota</th><th>Cuotas pagadas</th><th>Saldo</th><th>Estado</th></tr></thead>
                <tbody>
                  {prestamos.map((p) => (
                    <tr key={p.id}>
                      <td>{empleadoName(p.empleadoId)}</td>
                      <td>{p.fecha}</td>
                      <td style={{ fontSize: 12.5 }}>{p.motivo || '—'}</td>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(p.montoTotal)}</td>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(p.montoCuota)}</td>
                      <td>{p.cuotasPagadas} / {p.cuotas}</td>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace', color: p.saldoPendiente > 0 ? C.amber : C.cyanDark }}>{fmt(p.saldoPendiente)}</td>
                      <td>
                        <span className="text-xs font-semibold px-2 py-0.5 rounded-full" style={{ backgroundColor: p.estado === 'activo' ? '#FFF4E0' : '#E4F8F1', color: p.estado === 'activo' ? '#8A5A00' : C.cyanDark }}>
                          {p.estado === 'activo' ? 'Activo' : 'Liquidado'}
                        </span>
                      </td>
                    </tr>
                  ))}
                  {prestamos.length === 0 && <tr><td colSpan={8} style={{ color: C.muted, textAlign: 'center', padding: 18 }}>Sin préstamos registrados.</td></tr>}
                </tbody>
                {prestamos.length > 0 && (
                  <tfoot><tr><td colSpan={6} style={{ fontWeight: 700 }}>Saldo total pendiente</td><td colSpan={2} style={{ fontFamily: 'JetBrains Mono, monospace', fontWeight: 700, color: C.amber }}>{fmt(prestamos.reduce((s, p) => s + p.saldoPendiente, 0))}</td></tr></tfoot>
                )}
              </table>
            </div>

            {/* ORGANIGRAMA */}
            <div className="mt-8 pt-6" style={{ borderTop: `1px solid ${C.border}` }}>
              <h2 style={{ fontFamily: 'Space Grotesk, sans-serif', fontSize: 17, fontWeight: 700, color: C.slate }}>Organigrama</h2>
              <p style={{ color: C.muted, fontSize: 12.5, marginTop: 2 }}>Estructura jerárquica según a quién reporta cada empleado (editable en su tarjeta de arriba).</p>
              <Organigrama empleados={empleados} updateEmpleado={updateEmpleado} />
            </div>
          </div>
        )}

        {/* PLAN DE CUENTAS */}
        {tab === 'contab_plan' && (
          <div>
            <div className="flex items-center justify-between">
              <div>
                <h1 style={{ fontFamily: 'Space Grotesk, sans-serif', fontSize: 22, fontWeight: 700, color: C.slate }}>{t('Plan de cuentas')}</h1>
                <p style={{ color: C.muted, fontSize: 13, marginTop: 2 }}>Catálogo contable usado por todos los módulos. Aquí también configuras a qué línea del Estado de Resultados pertenece cada cuenta de Ingreso/Costo/Gasto.</p>
              </div>
              <button onClick={() => setShowCuentaForm(!showCuentaForm)} className="flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.navy }}>
                {showCuentaForm ? <X size={15} /> : <Plus size={15} />} {showCuentaForm ? 'Cerrar' : 'Nueva cuenta'}
              </button>
            </div>

            {showCuentaForm && <CuentaForm existentes={cuentas} onSubmit={(data) => { addCuenta(data); setShowCuentaForm(false); }} />}

            {cuentasSinGrupo.length > 0 && (
              <div className="mt-5 rounded-xl p-3.5 flex items-start gap-2.5" style={{ backgroundColor: '#FFF4E0', border: `1px solid ${C.amber}55` }}>
                <AlertTriangle size={16} color={C.amber} style={{ marginTop: 2, flexShrink: 0 }} />
                <div style={{ fontSize: 12.5, color: '#8A5A00' }}>
                  <strong>Sin agrupar:</strong> {cuentasSinGrupo.map((c) => c.nombre).join(', ')} no aparecerán en el Estado de Resultados hasta que les asignes un grupo abajo.
                </div>
              </div>
            )}

            <div className="rounded-xl overflow-hidden mt-5" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
              <table>
                <thead><tr><th>Código</th><th>Nombre (editable)</th><th>Tipo</th><th>Clasificación (Balance General)</th><th>Grupo (Estado de Resultados)</th><th></th></tr></thead>
                <tbody>
                  {cuentas.map((c) => (
                    <tr key={c.codigo}>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 12 }}>{c.codigo}</td>
                      <td><input value={c.nombre} onChange={(e) => updateCuentaNombre(c.codigo, e.target.value)} style={{ ...inputStyle, fontSize: 12.5, padding: '4px 6px', width: 240 }} /></td>
                      <td style={{ fontSize: 12.5 }}>{c.tipo}</td>
                      <td>
                        {['Activo', 'Pasivo'].includes(c.tipo) ? (
                          <select value={c.subtipo || ''} onChange={(e) => updateCuentaSubtipo(c.codigo, e.target.value)} style={{ ...inputStyle, fontSize: 12, padding: '4px 6px' }}>
                            <option value="">— Sin clasificar —</option>
                            {BG_SUBTIPOS.filter((s) => s.tipo === c.tipo).map((s) => <option key={s.key} value={s.key}>{s.label}</option>)}
                          </select>
                        ) : <span style={{ color: C.muted, fontSize: 12 }}>No aplica</span>}
                      </td>
                      <td>
                        {['Ingreso', 'Costo', 'Gasto'].includes(c.tipo) ? (
                          <select value={c.grupo || ''} onChange={(e) => updateCuentaGrupo(c.codigo, e.target.value)} style={{ ...inputStyle, fontSize: 12, padding: '4px 6px' }}>
                            <option value="">— Sin agrupar —</option>
                            {ER_GRUPOS.map((g) => <option key={g.key} value={g.key}>{g.label}</option>)}
                          </select>
                        ) : <span style={{ color: C.muted, fontSize: 12 }}>No aplica</span>}
                      </td>
                      <td><button onClick={() => deleteCuenta(c.codigo)}><Trash2 size={14} color={C.coral} /></button></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* COMPROBANTES (Libro diario) */}
        {tab === 'contab_comprobantes' && (
          <div>
            <div className="flex items-center justify-between">
              <div>
                <h1 style={{ fontFamily: 'Space Grotesk, sans-serif', fontSize: 22, fontWeight: 700, color: C.slate }}>{t('Comprobantes')}</h1>
                <p style={{ color: C.muted, fontSize: 13, marginTop: 2 }}>Libro diario: número de línea, código de cuenta, y el concepto debajo de cada asiento.</p>
              </div>
              <button onClick={() => { setShowAsientoForm(true); setEditingAsiento(null); }} className="flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.navy }}>
                <Plus size={15} /> Nuevo comprobante
              </button>
            </div>

            {(showAsientoForm || editingAsiento) && (
              <div style={{ position: 'fixed', inset: 0, backgroundColor: 'rgba(11,22,34,0.5)', zIndex: 70, display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 16 }}
                onClick={() => { setShowAsientoForm(false); setEditingAsiento(null); }}>
                <div onClick={(e) => e.stopPropagation()} className="rounded-xl p-5" style={{ backgroundColor: '#fff', width: 460, maxWidth: '100%', boxShadow: '0 12px 32px rgba(0,0,0,0.18)' }}>
                  <div style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 700, fontSize: 15, color: C.slate, marginBottom: 2 }}>
                    {editingAsiento ? `Editar comprobante ${editingAsiento.id}` : 'Nuevo comprobante'}
                  </div>
                  <AsientoForm key={editingAsiento?.id || 'nuevo'} cuentas={cuentas} centros={centros} initial={editingAsiento}
                    onSubmit={(data) => {
                      if (editingAsiento) { updateAsiento(editingAsiento.id, data); setEditingAsiento(null); }
                      else { addAsiento(data); setShowAsientoForm(false); }
                    }} />
                  <button onClick={() => { setShowAsientoForm(false); setEditingAsiento(null); }} className="w-full mt-2 py-2 rounded-lg text-sm font-semibold" style={{ backgroundColor: '#EEF1F4', color: C.slate }}>Cancelar</button>
                </div>
              </div>
            )}

            <div className="rounded-xl overflow-hidden mt-5" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
              <table>
                <thead><tr><th style={{ width: 32, textAlign: 'center' }}>#</th><th style={{ width: 80 }}>Fecha</th><th style={{ width: 70 }}>Código</th><th>Cuenta / concepto</th><th style={{ width: 50, textAlign: 'center' }}>Ref.</th><th style={{ width: 90, textAlign: 'right' }}>Debe</th><th style={{ width: 90, textAlign: 'right' }}>Haber</th><th>Usuario</th><th></th></tr></thead>
                <tbody>
                  {(() => {
                    let linea = comprobantesPag.inicio * 2 + 1;
                    return comprobantesPag.paginaActual.map((a) => {
                      const l1 = linea, l2 = linea + 1;
                      linea += 2;
                      const opacidad = a.anulado ? 0.5 : 1;
                      const tachado = a.anulado ? 'line-through' : 'none';
                      return (
                        <React.Fragment key={a.id}>
                          <tr style={{ borderTop: `2px solid ${C.border}` }}><td colSpan={9} style={{ textAlign: 'center', padding: 4, color: '#9AA5B1', fontSize: 11 }}>— {a.id} — {a.anulado && <span className="text-xs font-bold px-1.5 py-0.5 rounded" style={{ backgroundColor: '#FDE8E8', color: C.coral }}>ANULADO</span>}</td></tr>
                          <tr style={{ opacity: opacidad, textDecoration: tachado }}>
                            <td style={{ textAlign: 'center', color: '#9AA5B1' }}>{String(l1).padStart(2, '0')}</td>
                            <td>{a.fecha}</td>
                            <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{a.cuentaDebe}</td>
                            <td>{cuentaNombre(a.cuentaDebe)}</td>
                            <td rowSpan={4} style={{ textAlign: 'center', color: C.cyanDark, fontWeight: 600, verticalAlign: 'top', paddingTop: 8 }}>{a.id}</td>
                            <td style={{ textAlign: 'right', fontFamily: 'JetBrains Mono, monospace' }}>{fmt(a.montoDebe)}</td>
                            <td></td>
                            <td rowSpan={4} style={{ verticalAlign: 'top', color: C.muted, fontSize: 11.5 }}>{a.usuario || '—'}</td>
                            <td rowSpan={4} style={{ verticalAlign: 'top', paddingTop: 6, whiteSpace: 'nowrap' }}>
                              {!a.anulado && (
                                <div className="flex gap-2">
                                  <button onClick={() => { setEditingAsiento(a); setShowAsientoForm(false); }} title="Editar comprobante"><Pencil size={14} color={C.muted} /></button>
                                  <button onClick={() => anularAsiento(a.id)} title="Anular comprobante"><X size={15} color={C.coral} /></button>
                                </div>
                              )}
                            </td>
                          </tr>
                          <tr style={{ opacity: opacidad, textDecoration: tachado }}>
                            <td style={{ textAlign: 'center', color: '#9AA5B1' }}>{String(l2).padStart(2, '0')}</td>
                            <td></td>
                            <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{a.cuentaHaber}</td>
                            <td style={{ paddingLeft: 26, fontStyle: 'italic' }}>{cuentaNombre(a.cuentaHaber)}</td>
                            <td></td>
                            <td style={{ textAlign: 'right', fontFamily: 'JetBrains Mono, monospace' }}>{fmt(a.montoHaber)}</td>
                          </tr>
                          <tr style={{ opacity: opacidad }}>
                            <td></td><td></td><td></td>
                            <td style={{ paddingLeft: 26, color: C.muted, fontSize: 11.5, lineHeight: 1.5 }}>{a.glosa}</td>
                            <td></td><td></td>
                          </tr>
                          <tr style={{ opacity: opacidad }}>
                            <td></td><td></td><td></td>
                            <td style={{ paddingLeft: 26, color: C.muted, fontSize: 11.5 }}>{a.centroCostoId ? centroCostoName(a.centroCostoId) : ''}</td>
                            <td></td><td></td>
                          </tr>
                        </React.Fragment>
                      );
                    });
                  })()}
                </tbody>
              </table>
              <Paginador p={comprobantesPag} />
            </div>
          </div>
        )}

        {/* TESORERIA (Flujo de efectivo) */}
        {tab === 'contab_conciliacion' && (
          <div>
            <h1 style={{ fontFamily: 'Space Grotesk, sans-serif', fontSize: 22, fontWeight: 700, color: C.slate }}>{t('Conciliación bancaria')}</h1>
            <p style={{ color: C.muted, fontSize: 13, marginTop: 2 }}>Compara el saldo en libros de la cuenta Bancos (1110) contra tu estado de cuenta, marcando qué movimientos ya aparecen en el banco.</p>

            <div className="flex flex-wrap gap-3 items-end mt-5">
              <Field label="Fecha de corte"><input type="date" style={inputStyle} value={conciliacionCorte} onChange={(e) => setConciliacionCorte(e.target.value)} /></Field>
              <Field label="Saldo según banco ($)"><input type="number" style={{ ...inputStyle, width: 160 }} value={saldoBanco} onChange={(e) => setSaldoBanco(e.target.value)} placeholder="Del estado de cuenta" /></Field>
              <Field label="Importar estado de cuenta (CSV)">
                <input type="file" accept=".csv,.txt" onChange={(e) => { if (e.target.files[0]) importarEstadoCuenta(e.target.files[0]); e.target.value = ''; }} style={{ fontSize: 12 }} />
              </Field>
              {bancoMovs.length > 0 && (
                <button onClick={conciliarAutomatico} className="px-4 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.cyanDark }}>
                  Conciliar automáticamente ({bancoMovs.length} movs. del banco)
                </button>
              )}
            </div>
            <div style={{ fontSize: 11, color: C.muted, marginTop: 6 }}>Formato del CSV: una línea por movimiento con fecha, descripción, monto (negativo = cargo). Acepta fechas AAAA-MM-DD o DD/MM/AAAA, separadores coma o punto y coma. El matching automático empareja por monto exacto con hasta 7 días de diferencia.</div>

            {(() => {
              const movsCaja = asientosVigentes
                .filter((a) => (a.cuentaDebe === '1110' || a.cuentaHaber === '1110') && a.fecha <= conciliacionCorte)
                .sort((a, b) => b.fecha.localeCompare(a.fecha))
                .map((a) => ({ ...a, monto: a.cuentaDebe === '1110' ? a.montoDebe : -a.montoHaber }));
              const saldoLibros = movsCaja.reduce((s, a) => s + a.monto, 0);
              const noConciliados = movsCaja.filter((a) => !conciliados[a.id]);
              const totalNoConciliado = noConciliados.reduce((s, a) => s + a.monto, 0);
              const saldoConciliado = saldoLibros - totalNoConciliado;
              const banco = saldoBanco === '' ? null : Number(saldoBanco);
              const diferencia = banco === null ? null : banco - saldoConciliado;
              return (
                <>
                  <div className="grid grid-cols-4 gap-4 mt-5">
                    {[
                      ['Saldo en libros (1110 Bancos)', saldoLibros, C.slate],
                      ['(-) Movimientos sin conciliar', totalNoConciliado, C.amber],
                      ['= Saldo conciliado', saldoConciliado, C.cyanDark],
                      ['Diferencia vs banco', diferencia, diferencia === null ? C.muted : Math.abs(diferencia) < 0.01 ? C.cyanDark : C.coral],
                    ].map(([label, val, color], i) => (
                      <div key={i} className="rounded-xl p-4" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
                        <div style={{ fontSize: 11, fontWeight: 700, color: C.muted, textTransform: 'uppercase' }}>{label}</div>
                        <div style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 18, fontWeight: 700, color, marginTop: 6 }}>{val === null ? '—' : fmt(val)}</div>
                      </div>
                    ))}
                  </div>
                  {diferencia !== null && Math.abs(diferencia) < 0.01 && (
                    <div className="mt-3 rounded-xl p-3" style={{ backgroundColor: '#E4F8F1', border: `1px solid ${C.cyanDark}55`, fontSize: 12.5, color: '#0A6B5D' }}>✓ Conciliado: el saldo del banco coincide con los movimientos marcados.</div>
                  )}
                  <div className="rounded-xl overflow-hidden mt-5" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
                    <table>
                      <thead><tr><th>Conciliado</th><th>Fecha</th><th>Glosa</th><th>Entrada</th><th>Salida</th></tr></thead>
                      <tbody>
                        {movsCaja.map((a) => (
                          <tr key={a.id} style={{ opacity: conciliados[a.id] ? 1 : 0.75 }}>
                            <td><input type="checkbox" checked={!!conciliados[a.id]} onChange={() => setConciliados((prev) => ({ ...prev, [a.id]: !prev[a.id] }))} style={{ width: 16, height: 16, accentColor: C.cyanDark }} /></td>
                            <td>{a.fecha}</td>
                            <td style={{ fontSize: 12.5 }}>{a.glosa}</td>
                            <td style={{ fontFamily: 'JetBrains Mono, monospace', color: C.cyanDark }}>{a.monto > 0 ? fmt(a.monto) : ''}</td>
                            <td style={{ fontFamily: 'JetBrains Mono, monospace', color: C.coral }}>{a.monto < 0 ? fmt(-a.monto) : ''}</td>
                          </tr>
                        ))}
                        {movsCaja.length === 0 && <tr><td colSpan={5} style={{ color: C.muted, textAlign: 'center', padding: 20 }}>Sin movimientos de caja a esta fecha.</td></tr>}
                      </tbody>
                    </table>
                  </div>

                  {bancoMovs.length > 0 && (
                    <div className="rounded-xl overflow-hidden mt-5" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
                      <div className="px-4 pt-4 pb-1" style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 600, fontSize: 14, color: C.slate }}>Estado de cuenta importado</div>
                      <table>
                        <thead><tr><th>Fecha</th><th>Descripción (banco)</th><th>Monto</th><th>Resultado del matching</th></tr></thead>
                        <tbody>
                          {bancoMovs.map((b) => (
                            <tr key={b.id}>
                              <td>{b.fecha}</td>
                              <td style={{ fontSize: 12.5 }}>{b.descripcion}</td>
                              <td style={{ fontFamily: 'JetBrains Mono, monospace', color: b.monto >= 0 ? C.cyanDark : C.coral }}>{fmt(b.monto)}</td>
                              <td>
                                {b.estado === 'conciliado' && <span className="text-xs font-semibold px-2 py-0.5 rounded-full" style={{ backgroundColor: '#E4F8F1', color: C.cyanDark }}>✓ Conciliado con {b.asientoId}</span>}
                                {b.estado === 'sin_coincidencia' && <span className="text-xs font-semibold px-2 py-0.5 rounded-full" style={{ backgroundColor: '#FDE8E8', color: C.coral }}>Sin coincidencia en libros</span>}
                                {b.estado === 'pendiente' && <span className="text-xs font-semibold px-2 py-0.5 rounded-full" style={{ backgroundColor: '#FFF4E0', color: C.amber }}>Pendiente — presiona "Conciliar automáticamente"</span>}
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                      <div style={{ fontSize: 11.5, color: C.muted, padding: '8px 16px 14px' }}>Los movimientos "sin coincidencia" pueden ser comisiones bancarias o intereses no registrados aún — regístralos como comprobante y vuelve a conciliar.</div>
                    </div>
                  )}

                  <div className="flex items-center gap-3 mt-5">
                    <button onClick={() => guardarConciliacion({ corte: conciliacionCorte, saldoBanco: banco, saldoLibros, saldoConciliado, diferencia, movimientosConciliados: movsCaja.filter((a) => conciliados[a.id]).length, movimientosTotales: movsCaja.length })}
                      className="px-4 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.navy }}>
                      Guardar conciliación
                    </button>
                    <span style={{ fontSize: 12, color: C.muted }}>Guarda un registro con el corte, saldos y movimientos conciliados de esta sesión.</span>
                  </div>

                  {conciliacionesGuardadas.length > 0 && (
                    <div className="rounded-xl overflow-hidden mt-5" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
                      <div className="px-4 pt-4 pb-1" style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 600, fontSize: 14, color: C.slate }}>Historial de conciliaciones guardadas</div>
                      <table>
                        <thead><tr><th>ID</th><th>Guardada el</th><th>Fecha de corte</th><th>Saldo banco</th><th>Saldo libros</th><th>Conciliados</th><th>Diferencia</th></tr></thead>
                        <tbody>
                          {[...conciliacionesGuardadas].reverse().map((cn) => (
                            <tr key={cn.id}>
                              <td style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 12 }}>{cn.id}</td>
                              <td style={{ fontSize: 12 }}>{cn.guardadoEl}</td>
                              <td>{cn.corte}</td>
                              <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{cn.saldoBanco === null ? '—' : fmt(cn.saldoBanco)}</td>
                              <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(cn.saldoLibros)}</td>
                              <td style={{ fontSize: 12 }}>{cn.movimientosConciliados} / {cn.movimientosTotales}</td>
                              <td style={{ fontFamily: 'JetBrains Mono, monospace', color: cn.diferencia === null ? C.muted : Math.abs(cn.diferencia) < 0.01 ? C.cyanDark : C.coral }}>{cn.diferencia === null ? '—' : fmt(cn.diferencia)}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  )}
                </>
              );
            })()}
          </div>
        )}

        {/* INFORMES (Balance general, Estado de resultados, Balance de comprobación, Libro mayor) */}
        {tab === 'contab_informes' && (
          <div>
            <h1 style={{ fontFamily: 'Space Grotesk, sans-serif', fontSize: 22, fontWeight: 700, color: C.slate }}>{t('Informes')}</h1>
            <p style={{ color: C.muted, fontSize: 13, marginTop: 2 }}>Elige el informe contable que necesitas.</p>

            <div className="flex flex-wrap gap-3 items-end mt-5">
              <Field label="Informe">
                <select style={{ ...inputStyle, width: 240 }} value={contabVista} onChange={(e) => setContabVista(e.target.value)}>
                  {CONTAB_VISTAS.map((v) => <option key={v.key} value={v.key}>{v.label}</option>)}
                </select>
              </Field>
              {contabVista === 'estado_resultados' && (
                <Field label="Centro de costo">
                  <select style={{ ...inputStyle, width: 220 }} value={estadoResCentro} onChange={(e) => setEstadoResCentro(e.target.value)}>
                    <option value="todos">Todos los centros (consolidado)</option>
                    {centros.map((c) => <option key={c.id} value={c.id}>{c.nombre}</option>)}
                  </select>
                </Field>
              )}
              {contabVista === 'estado_resultados' && estadoResCentro !== 'todos' && estadoResCentro !== 'CC-ADM' && (
                <Field label="Prorrateo de gastos admin">
                  <select style={{ ...inputStyle, width: 180 }} value={prorrateoMetodo} onChange={(e) => setProrrateoMetodo(e.target.value)}>
                    <option value="ventas">En función de ventas</option>
                    <option value="equitativo">Equitativo (partes iguales)</option>
                  </select>
                </Field>
              )}
              {contabVista === 'estado_resultados' && (
                <Field label="Modo">
                  <select style={{ ...inputStyle, width: 200 }} value={erModo} onChange={(e) => setErModo(e.target.value)}>
                    <option value="periodo">Del periodo (Del/Al)</option>
                    <option value="acumulado">Acumulado (hasta "Al")</option>
                    <option value="comparativo">Comparativo (A vs B)</option>
                  </select>
                </Field>
              )}
              {['estado_resultados', 'flujo_efectivo'].includes(contabVista) && (
                <>
                  <Field label={contabVista === 'estado_resultados' && erModo === 'comparativo' ? 'Periodo A — Del' : contabVista === 'estado_resultados' && erModo === 'acumulado' ? 'Del (ignorado en acumulado)' : 'Del'}><input type="date" style={inputStyle} value={reportFrom} onChange={(e) => setReportFrom(e.target.value)} /></Field>
                  <Field label={contabVista === 'estado_resultados' && erModo === 'comparativo' ? 'Periodo A — Al' : 'Al'}><input type="date" style={inputStyle} value={reportTo} onChange={(e) => setReportTo(e.target.value)} /></Field>
                </>
              )}
              {contabVista === 'estado_resultados' && erModo === 'comparativo' && (
                <>
                  <Field label="Periodo B — Del"><input type="date" style={inputStyle} value={compareFrom} onChange={(e) => setCompareFrom(e.target.value)} /></Field>
                  <Field label="Periodo B — Al"><input type="date" style={inputStyle} value={compareTo} onChange={(e) => setCompareTo(e.target.value)} /></Field>
                </>
              )}
              {contabVista === 'balance_general' && (
                <Field label="Al (fecha de corte)"><input type="date" style={inputStyle} value={reportTo} onChange={(e) => setReportTo(e.target.value)} /></Field>
              )}
              <button onClick={exportInformeActual} className="flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-white text-sm font-semibold no-print" style={{ backgroundColor: C.navy }}>
                <Download size={14} /> Exportar CSV
              </button>
              <button onClick={() => window.print()} className="flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-sm font-semibold no-print" style={{ backgroundColor: '#fff', color: C.slate, border: `1px solid ${C.border}` }}>
                Imprimir / PDF
              </button>
            </div>

            {/* ESTADO DE RESULTADOS COMPARATIVO */}
            {contabVista === 'estado_resultados' && erModo === 'comparativo' && (() => {
              const A = erTotalsFor(reportFrom, reportTo, estadoResCentro);
              const B = erTotalsFor(compareFrom, compareTo, estadoResCentro);
              const filas = [
                ['Ingresos por instalaciones', A.ingresos, B.ingresos, false],
                ['(-) Costo de materiales', -A.cMat, -B.cMat, false],
                ['(-) Costo de mano de obra', -A.cMO, -B.cMO, false],
                ['(-) Costo de depreciación', -A.cDep, -B.cDep, false],
                ['= Utilidad operativa', A.utilOp, B.utilOp, true],
                ['(+) Otros ingresos', A.otros, B.otros, false],
                ['(-) Gastos generales', -A.gastos, -B.gastos, false],
                ['= Utilidad neta', A.utilNeta, B.utilNeta, true],
              ];
              return (
                <div className="rounded-xl overflow-hidden mt-6" style={{ backgroundColor: C.card, border: `1px solid ${C.border}`, maxWidth: 820 }}>
                  <table>
                    <thead><tr><th>Concepto</th><th>A: {reportFrom} → {reportTo}</th><th>B: {compareFrom} → {compareTo}</th><th>Variación $</th><th>Variación %</th></tr></thead>
                    <tbody>
                      {filas.map(([label, va, vb, strong], i) => {
                        const dif = va - vb;
                        const pct = vb !== 0 ? (dif / Math.abs(vb)) * 100 : null;
                        return (
                          <tr key={i} style={{ fontWeight: strong ? 700 : 400, backgroundColor: strong ? '#F6F8FA' : 'transparent' }}>
                            <td style={{ fontSize: 13 }}>{label}</td>
                            <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(va)}</td>
                            <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(vb)}</td>
                            <td style={{ fontFamily: 'JetBrains Mono, monospace', color: dif >= 0 ? C.cyanDark : C.coral }}>{dif >= 0 ? '+' : ''}{fmt(dif)}</td>
                            <td style={{ fontFamily: 'JetBrains Mono, monospace', color: dif >= 0 ? C.cyanDark : C.coral }}>{pct === null ? '—' : (pct >= 0 ? '+' : '') + pct.toFixed(1) + '%'}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                  <div style={{ fontSize: 11.5, color: C.muted, padding: '8px 16px 14px' }}>Usa Periodo A = mes actual y Periodo B = mes anterior (o mismo mes del año pasado) para ver la variación mensual o anual.</div>
                </div>
              );
            })()}

            {/* ESTADO DE RESULTADOS ACUMULADO */}
            {contabVista === 'estado_resultados' && erModo === 'acumulado' && (() => {
              const T = erTotalsFor('', reportTo, estadoResCentro);
              return (
                <div className="rounded-xl p-5 mt-6" style={{ backgroundColor: C.card, border: `1px solid ${C.border}`, maxWidth: 560 }}>
                  <div style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 600, fontSize: 14, color: C.slate, marginBottom: 10 }}>Estado de resultados acumulado — desde el inicio hasta {reportTo}</div>
                  {[
                    ['Ingresos por instalaciones', T.ingresos, false],
                    ['(-) Costo de materiales', -T.cMat, false],
                    ['(-) Costo de mano de obra', -T.cMO, false],
                    ['(-) Costo de depreciación', -T.cDep, false],
                    ['= Utilidad operativa', T.utilOp, true],
                    ['(+) Otros ingresos', T.otros, false],
                    ['(-) Gastos generales', -T.gastos, false],
                    ['= Utilidad neta acumulada', T.utilNeta, true],
                  ].map(([label, val, strong], i) => (
                    <div key={i} className="flex justify-between py-1.5" style={{ borderTop: strong ? `1px solid ${C.border}` : 'none', marginTop: strong ? 6 : 0 }}>
                      <span style={{ fontSize: 13, fontWeight: strong ? 700 : 400, color: C.slate }}>{label}</span>
                      <span style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 13, fontWeight: strong ? 700 : 500, color: val >= 0 ? (strong ? C.cyanDark : C.slate) : C.coral }}>{fmt(val)}</span>
                    </div>
                  ))}
                </div>
              );
            })()}

            {/* FLUJO DE EFECTIVO */}
            {contabVista === 'flujo_efectivo' && (
              <div className="grid grid-cols-2 gap-5 mt-6">
                <div className="rounded-xl p-5" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
                  <div style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 600, fontSize: 14, color: C.slate, marginBottom: 10 }}>Estado de flujo de efectivo</div>
                  {[
                    ['Saldo inicial de efectivo', saldoInicialCaja, false],
                    ['Flujo de actividades de operación', flujoOperacion, false],
                    ['Flujo de actividades de inversión', flujoInversion, false],
                    ['Flujo de actividades de financiamiento', flujoFinanciamiento, false],
                    ['= Flujo neto del periodo', flujoNetoPeriodo, true],
                    ['Saldo final de efectivo', saldoFinalCaja, true],
                  ].map(([label, val, strong], i) => (
                    <div key={i} className="flex justify-between py-1.5" style={{ borderTop: strong ? `1px solid ${C.border}` : 'none', marginTop: strong ? 6 : 0 }}>
                      <span style={{ fontSize: 13, fontWeight: strong ? 700 : 400, color: C.slate }}>{label}</span>
                      <span style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 13, fontWeight: strong ? 700 : 500, color: val >= 0 ? (strong ? C.cyanDark : C.slate) : C.coral }}>{fmt(val)}</span>
                    </div>
                  ))}
                </div>

                <div className="rounded-xl overflow-hidden" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
                  <div className="px-4 pt-4 pb-1" style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 600, fontSize: 14, color: C.slate }}>Detalle por actividad</div>
                  {[['Operación', flujoDetalle.operacion], ['Inversión', flujoDetalle.inversion], ['Financiamiento', flujoDetalle.financiamiento]].map(([label, list]) => (
                    <div key={label} className="px-4 pb-3">
                      <div style={{ fontSize: 11.5, fontWeight: 700, color: C.muted, textTransform: 'uppercase', marginTop: 10, marginBottom: 4 }}>{label}</div>
                      {list.length === 0 && <div style={{ fontSize: 12.5, color: C.muted }}>Sin movimientos.</div>}
                      {list.map((a) => (
                        <div key={a.id} className="flex justify-between py-0.5" style={{ fontSize: 12.5 }}>
                          <span style={{ color: C.slate }}>{a.glosa}</span>
                          <span style={{ fontFamily: 'JetBrains Mono, monospace', color: a.monto >= 0 ? C.cyanDark : C.coral }}>{fmt(a.monto)}</span>
                        </div>
                      ))}
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* BALANCE GENERAL */}
            {contabVista === 'balance_general' && (() => {
              const seccion = (lista, subtipo) => lista.filter((c) => c.subtipo === subtipo);
              const sinClasificarA = cuentasActivo.filter((c) => !c.subtipo);
              const sinClasificarP = cuentasPasivo.filter((c) => !c.subtipo);
              const subtotal = (l) => l.reduce((s, c) => s + c.saldo, 0);
              const Bloque = ({ titulo, lista }) => (
                <div style={{ marginBottom: 12 }}>
                  <div style={{ fontSize: 12, fontWeight: 700, color: C.muted, textTransform: 'uppercase', marginBottom: 4 }}>{titulo}</div>
                  {lista.map((c) => (
                    <div key={c.codigo} className="flex justify-between py-1"><span style={{ fontSize: 13, color: C.slate }}>{c.nombre}</span><span style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 13, color: c.saldo < 0 ? C.coral : C.slate }}>{fmt(c.saldo)}</span></div>
                  ))}
                  {lista.length === 0 && <div style={{ fontSize: 12, color: C.muted }}>Sin cuentas.</div>}
                  <div className="flex justify-between py-1" style={{ borderTop: `1px dashed ${C.border}` }}>
                    <span style={{ fontSize: 12.5, fontWeight: 600, color: C.slate }}>Subtotal {titulo.toLowerCase()}</span>
                    <span style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 12.5, fontWeight: 600 }}>{fmt(subtotal(lista))}</span>
                  </div>
                </div>
              );
              return (
              <div className="grid grid-cols-2 gap-5 mt-6">
                <div className="rounded-xl p-5" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
                  <div style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 600, fontSize: 14, color: C.slate, marginBottom: 10 }}>Activo</div>
                  <Bloque titulo="Activo corriente" lista={seccion(cuentasActivo, 'activo_corriente')} />
                  <Bloque titulo="Activo no corriente" lista={seccion(cuentasActivo, 'activo_no_corriente')} />
                  {sinClasificarA.length > 0 && <Bloque titulo="Sin clasificar" lista={sinClasificarA} />}
                  <div className="flex justify-between py-1.5 mt-1" style={{ borderTop: `2px solid ${C.navy}` }}>
                    <span style={{ fontSize: 13, fontWeight: 700, color: C.slate }}>Total activo</span>
                    <span style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 13, fontWeight: 700, color: C.cyanDark }}>{fmt(totalActivo)}</span>
                  </div>
                </div>

                <div className="rounded-xl p-5" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
                  <div style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 600, fontSize: 14, color: C.slate, marginBottom: 10 }}>Pasivo</div>
                  <Bloque titulo="Pasivo corriente" lista={seccion(cuentasPasivo, 'pasivo_corriente')} />
                  <Bloque titulo="Pasivo no corriente" lista={seccion(cuentasPasivo, 'pasivo_no_corriente')} />
                  {sinClasificarP.length > 0 && <Bloque titulo="Sin clasificar" lista={sinClasificarP} />}
                  <div className="flex justify-between py-1.5 mt-1" style={{ borderTop: `1px solid ${C.border}` }}>
                    <span style={{ fontSize: 13, fontWeight: 700, color: C.slate }}>Total pasivo</span>
                    <span style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 13, fontWeight: 700 }}>{fmt(totalPasivo)}</span>
                  </div>

                  <div style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 600, fontSize: 14, color: C.slate, margin: '18px 0 10px' }}>Capital</div>
                  {cuentasCapital.map((c) => (
                    <div key={c.codigo} className="flex justify-between py-1"><span style={{ fontSize: 13, color: C.slate }}>{c.nombre}</span><span style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 13 }}>{fmt(c.saldo)}</span></div>
                  ))}
                  <div className="flex justify-between py-1"><span style={{ fontSize: 13, color: C.slate }}>Utilidad del ejercicio</span><span style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 13, color: utilidadEjercicio >= 0 ? C.cyanDark : C.coral }}>{fmt(utilidadEjercicio)}</span></div>
                  <div className="flex justify-between py-1.5 mt-1" style={{ borderTop: `1px solid ${C.border}` }}>
                    <span style={{ fontSize: 13, fontWeight: 700, color: C.slate }}>Total capital</span>
                    <span style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 13, fontWeight: 700 }}>{fmt(totalCapital)}</span>
                  </div>

                  <div className="flex justify-between py-1.5 mt-3" style={{ borderTop: `2px solid ${C.navy}` }}>
                    <span style={{ fontSize: 13, fontWeight: 700, color: C.slate }}>Total pasivo + capital</span>
                    <span style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 13, fontWeight: 700, color: totalPasivoCapital === totalActivo ? C.cyanDark : C.coral }}>{fmt(totalPasivoCapital)}</span>
                  </div>
                </div>
              </div>
              );
            })()}

            {/* ESTADO DE RESULTADOS */}
            {contabVista === 'estado_resultados' && erModo === 'periodo' && (
              <div className="rounded-xl p-5 mt-6" style={{ backgroundColor: C.card, border: `1px solid ${C.border}`, maxWidth: 620 }}>
                <div className="flex items-center justify-between" style={{ marginBottom: 4 }}>
                  <div style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 600, fontSize: 14, color: C.slate }}>
                    {estadoResCentro === 'todos' ? 'Estado de resultados — consolidado' : `Estado de resultados — ${centroCostoName(estadoResCentro)}`}
                  </div>
                  <label className="flex items-center gap-1.5" style={{ fontSize: 11.5, color: C.muted, cursor: 'pointer' }}>
                    <input type="checkbox" checked={erMostrarDetalle} onChange={(e) => setErMostrarDetalle(e.target.checked)} style={{ width: 13, height: 13, accentColor: C.cyanDark }} />
                    Ver detalle por cuenta
                  </label>
                </div>
                <div style={{ fontSize: 11.5, color: C.muted, marginBottom: 10 }}>Calculado directamente del libro mayor, agrupado según el Plan de Cuentas.</div>

                {erLineas.map((linea, i) => (
                  <div key={i} style={{ marginTop: 8 }}>
                    <div className="flex justify-between py-1">
                      <span style={{ fontSize: 13, color: C.slate }}>{linea.label}</span>
                      <span style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 13, color: linea.total >= 0 ? C.slate : C.coral }}>{fmt(linea.total)}</span>
                    </div>
                    {erMostrarDetalle && linea.cuentas.length > 0 && (
                      <div style={{ marginLeft: 14, borderLeft: `2px solid ${C.border}`, paddingLeft: 10 }}>
                        {linea.cuentas.map((c) => (
                          <div key={c.codigo} className="flex justify-between py-0.5">
                            <span style={{ fontSize: 11.5, color: C.muted }}>{c.codigo} — {c.nombre}</span>
                            <span style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 11.5, color: C.muted }}>{fmt(linea.resta ? -c.saldo : c.saldo)}</span>
                          </div>
                        ))}
                      </div>
                    )}
                    {erMostrarDetalle && linea.cuentas.length === 0 && <div style={{ marginLeft: 14, fontSize: 11.5, color: C.amber }}>Sin cuentas asignadas a este grupo — ve a Plan de Cuentas para configurarlo.</div>}
                  </div>
                ))}

                <div className="flex justify-between py-1.5 mt-2" style={{ borderTop: `1px solid ${C.border}` }}>
                  <span style={{ fontSize: 13, fontWeight: 700, color: C.slate }}>= Utilidad operativa</span>
                  <span style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 13, fontWeight: 700, color: erUtilidadOperativa >= 0 ? C.cyanDark : C.coral }}>{fmt(erUtilidadOperativa)}</span>
                </div>

                {erLineasCorporativas.map((linea, i) => (
                  <div key={i} style={{ marginTop: 8 }}>
                    <div className="flex justify-between py-1">
                      <span style={{ fontSize: 13, color: C.slate }}>{linea.label}</span>
                      <span style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 13, color: linea.total >= 0 ? C.slate : C.coral }}>{fmt(linea.total)}</span>
                    </div>
                    {erMostrarDetalle && linea.cuentas.length > 0 && (
                      <div style={{ marginLeft: 14, borderLeft: `2px solid ${C.border}`, paddingLeft: 10 }}>
                        {linea.cuentas.map((c) => (
                          <div key={c.codigo} className="flex justify-between py-0.5">
                            <span style={{ fontSize: 11.5, color: C.muted }}>{c.codigo} — {c.nombre}</span>
                            <span style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 11.5, color: C.muted }}>{fmt(linea.resta ? -c.saldo : c.saldo)}</span>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                ))}
                <div className="flex justify-between py-1.5 mt-2" style={{ borderTop: `2px solid ${C.navy}` }}>
                  <span style={{ fontSize: 13, fontWeight: 700, color: C.slate }}>= Utilidad neta del periodo</span>
                  <span style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 13, fontWeight: 700, color: erUtilidadNeta >= 0 ? C.cyanDark : C.coral }}>{fmt(erUtilidadNeta)}</span>
                </div>
                {esCentroOperativo && erProrrateoInfo.monto > 0 && (
                  <div style={{ fontSize: 11, color: C.muted, marginTop: 8 }}>La línea de Gastos generales incluye {fmt(erProrrateoInfo.monto)} de gastos administrativos prorrateados ({prorrateoMetodo === 'ventas' ? 'en proporción a los ingresos por órdenes del centro' : 'en partes iguales'}). La vista consolidada muestra el total real sin prorratear.</div>
                )}
              </div>
            )}

            {/* BALANCE DE COMPROBACION */}
            {contabVista === 'balance_comprobacion' && (
              <div className="rounded-xl overflow-hidden mt-6" style={{ backgroundColor: C.card, border: `1px solid ${C.border}`, maxWidth: 640 }}>
                <table>
                  <thead><tr><th>Cuenta</th><th>Debe</th><th>Haber</th></tr></thead>
                  <tbody>
                    {balanceComprobacion.map((c) => (
                      <tr key={c.codigo}>
                        <td style={{ fontSize: 12.5 }}>{c.codigo} — {c.nombre}</td>
                        <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(c.debe)}</td>
                        <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(c.haber)}</td>
                      </tr>
                    ))}
                  </tbody>
                  <tfoot>
                    <tr>
                      <td style={{ fontWeight: 700 }}>Totales</td>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace', fontWeight: 700 }}>{fmt(totalDebe)}</td>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace', fontWeight: 700 }}>{fmt(totalHaber)}</td>
                    </tr>
                  </tfoot>
                </table>
              </div>
            )}

            {/* LIBRO MAYOR */}
            {contabVista === 'libro_mayor' && (
              <div className="rounded-xl overflow-hidden mt-6" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
                <div className="px-4 pt-4 pb-2 flex items-center justify-between flex-wrap gap-2">
                  <span style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 600, fontSize: 14, color: C.slate }}>Libro mayor</span>
                  <div className="flex gap-2">
                    <select value={mayorCentro} onChange={(e) => setMayorCentro(e.target.value)} style={{ ...inputStyle, fontSize: 12, padding: '4px 8px' }}>
                      <option value="todos">Todos los centros</option>
                      {centros.map((c) => <option key={c.id} value={c.id}>{c.nombre}</option>)}
                    </select>
                    <select value={cuentaMayor} onChange={(e) => setCuentaMayor(e.target.value)} style={{ ...inputStyle, fontSize: 12, padding: '4px 8px' }}>
                      {cuentas.map((c) => <option key={c.codigo} value={c.codigo}>{c.codigo} — {c.nombre}</option>)}
                    </select>
                  </div>
                </div>
                <table>
                  <thead><tr><th>Fecha</th><th>Glosa</th><th>Debe</th><th>Haber</th><th>Saldo</th></tr></thead>
                  <tbody>
                    {mayorRows.map((r) => (
                      <tr key={r.id}>
                        <td>{r.fecha}</td>
                        <td style={{ fontSize: 12 }}>{r.glosa}</td>
                        <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{r.debe ? fmt(r.debe) : ''}</td>
                        <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{r.haber ? fmt(r.haber) : ''}</td>
                        <td style={{ fontFamily: 'JetBrains Mono, monospace', fontWeight: 600 }}>{fmt(r.saldo)}</td>
                      </tr>
                    ))}
                    {mayorRows.length === 0 && <tr><td colSpan={5} style={{ color: C.muted, textAlign: 'center', padding: 20 }}>Sin movimientos.</td></tr>}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        )}

        {/* CONFIGURACIÓN DEL MÓDULO CONTABLE */}
        {tab === 'contab_config' && (
          <div>
            <h1 style={{ fontFamily: 'Space Grotesk, sans-serif', fontSize: 22, fontWeight: 700, color: C.slate }}>{t('Configuración')}</h1>
            <p style={{ color: C.muted, fontSize: 13, marginTop: 2 }}>Parametriza los catálogos que usa todo el módulo contable: elige qué configurar.</p>

            <div className="mt-5">
              <Field label="Tipo de configuración">
                <select style={{ ...inputStyle, width: 260 }} value={configSeccion} onChange={(e) => setConfigSeccion(e.target.value)}>
                  {CONFIG_SECCIONES.map((s) => <option key={s.key} value={s.key}>{s.label}</option>)}
                </select>
              </Field>
            </div>

            {configSeccion === 'centros' && (<>
            <div className="flex items-center justify-between mt-7">
              <div style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 600, fontSize: 15, color: C.slate }}>Centros de costo</div>
              <button onClick={() => setShowCentroForm(!showCentroForm)} className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-white text-xs font-semibold" style={{ backgroundColor: C.navy }}>
                {showCentroForm ? <X size={13} /> : <Plus size={13} />} {showCentroForm ? 'Cerrar' : 'Nuevo centro'}
              </button>
            </div>
            {showCentroForm && <CentroForm onSubmit={(data) => { addCentro(data); setShowCentroForm(false); }} />}
            <div className="rounded-xl overflow-hidden mt-3" style={{ backgroundColor: C.card, border: `1px solid ${C.border}`, maxWidth: 520 }}>
              <table>
                <thead><tr><th>Código</th><th>Nombre</th><th></th></tr></thead>
                <tbody>
                  {centros.map((c) => (
                    <tr key={c.id}>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 12 }}>{c.codigo}</td>
                      <td>{c.nombre}</td>
                      <td><button onClick={() => deleteCentro(c.id)}><Trash2 size={14} color={C.coral} /></button></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            </>)}

            {configSeccion === 'cat_gasto' && (<>
            <div className="flex items-center justify-between mt-7">
              <div style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 600, fontSize: 15, color: C.slate }}>Categorías de gasto</div>
              <button onClick={() => setShowCatGastoForm(!showCatGastoForm)} className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-white text-xs font-semibold" style={{ backgroundColor: C.navy }}>
                {showCatGastoForm ? <X size={13} /> : <Plus size={13} />} {showCatGastoForm ? 'Cerrar' : 'Nueva categoría'}
              </button>
            </div>
            {showCatGastoForm && <CategoriaForm cuentas={cuentas.filter((c) => c.tipo === 'Gasto' || c.tipo === 'Costo')} onSubmit={(data) => { addCategoria('gasto', data); setShowCatGastoForm(false); }} />}
            <div className="rounded-xl overflow-hidden mt-3" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
              <table>
                <thead><tr><th>Categoría</th><th>Cuenta contable asociada</th><th></th></tr></thead>
                <tbody>
                  {gastoCategorias.map((c) => (
                    <tr key={c.nombre}>
                      <td>{c.nombre}</td>
                      <td>
                        <select value={c.cuenta} onChange={(e) => updateCategoriaCuenta('gasto', c.nombre, e.target.value)} style={{ ...inputStyle, fontSize: 12, padding: '4px 6px' }}>
                          {cuentas.filter((cta) => cta.tipo === 'Gasto' || cta.tipo === 'Costo').map((cta) => <option key={cta.codigo} value={cta.codigo}>{cta.codigo} — {cta.nombre}</option>)}
                        </select>
                      </td>
                      <td><button onClick={() => deleteCategoria('gasto', c.nombre)}><Trash2 size={14} color={C.coral} /></button></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            </>)}

            {configSeccion === 'cat_ingreso' && (<>
            <div className="flex items-center justify-between mt-7">
              <div style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 600, fontSize: 15, color: C.slate }}>Categorías de ingreso</div>
              <button onClick={() => setShowCatIngresoForm(!showCatIngresoForm)} className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-white text-xs font-semibold" style={{ backgroundColor: C.navy }}>
                {showCatIngresoForm ? <X size={13} /> : <Plus size={13} />} {showCatIngresoForm ? 'Cerrar' : 'Nueva categoría'}
              </button>
            </div>
            {showCatIngresoForm && <CategoriaForm cuentas={cuentas.filter((c) => c.tipo === 'Ingreso')} onSubmit={(data) => { addCategoria('ingreso', data); setShowCatIngresoForm(false); }} />}
            <div className="rounded-xl overflow-hidden mt-3" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
              <table>
                <thead><tr><th>Categoría</th><th>Cuenta contable asociada</th><th></th></tr></thead>
                <tbody>
                  {ingresoCategorias.map((c) => (
                    <tr key={c.nombre}>
                      <td>{c.nombre}</td>
                      <td>
                        <select value={c.cuenta} onChange={(e) => updateCategoriaCuenta('ingreso', c.nombre, e.target.value)} style={{ ...inputStyle, fontSize: 12, padding: '4px 6px' }}>
                          {cuentas.filter((cta) => cta.tipo === 'Ingreso').map((cta) => <option key={cta.codigo} value={cta.codigo}>{cta.codigo} — {cta.nombre}</option>)}
                        </select>
                      </td>
                      <td><button onClick={() => deleteCategoria('ingreso', c.nombre)}><Trash2 size={14} color={C.coral} /></button></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            </>)}

            {configSeccion === 'partidas' && (<>
            <div style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 600, fontSize: 15, color: C.slate, marginTop: 28 }}>Partidas de trabajo / servicios</div>
            <p style={{ fontSize: 12, color: C.muted, marginTop: 2 }}>Adapta este catálogo al giro de tu empresa — son las líneas que se pueden desglosar al crear una orden (mano de obra, materiales, servicios específicos, etc.).</p>
            <div className="rounded-xl overflow-hidden mt-3" style={{ backgroundColor: C.card, border: `1px solid ${C.border}`, maxWidth: 520 }}>
              <table>
                <thead><tr><th>Partida</th><th>Precio de referencia</th><th></th></tr></thead>
                <tbody>
                  {catalogoPartidas.map((p) => (
                    <tr key={p.id}>
                      <td><input value={p.nombre} onChange={(e) => { setCatalogoPartidas((prev) => prev.map((x) => x.id === p.id ? { ...x, nombre: e.target.value } : x)); supabase.from('catalogo_partidas').update({ nombre: e.target.value }).eq('id', p.id).then(() => {}); }} style={{ ...inputStyle, fontSize: 12.5, padding: '4px 6px', width: 220 }} /></td>
                      <td><input type="number" value={p.precioRef} onChange={(e) => { const v = Number(e.target.value) || 0; setCatalogoPartidas((prev) => prev.map((x) => x.id === p.id ? { ...x, precioRef: v } : x)); supabase.from('catalogo_partidas').update({ precio_ref: v }).eq('id', p.id).then(() => {}); }} style={{ ...inputStyle, fontSize: 12.5, padding: '4px 6px', width: 100, fontFamily: 'JetBrains Mono, monospace' }} /></td>
                      <td><button onClick={() => { setCatalogoPartidas((prev) => prev.filter((x) => x.id !== p.id)); supabase.from('catalogo_partidas').delete().eq('id', p.id).then(() => {}); }}><Trash2 size={14} color={C.coral} /></button></td>
                    </tr>
                  ))}
                  {catalogoPartidas.length === 0 && <tr><td colSpan={3} style={{ color: C.muted, textAlign: 'center', padding: 16 }}>Sin partidas configuradas.</td></tr>}
                </tbody>
              </table>
              <div className="p-3">
                <button onClick={async () => {
                  const { data: nueva, error } = await supabase.from('catalogo_partidas').insert({ empresa_id: empresaActiva, nombre: 'Nueva partida', precio_ref: 0 }).select().single();
                  if (error) { alert('No se pudo guardar la partida: ' + error.message); return; }
                  setCatalogoPartidas((prev) => [...prev, { id: nueva.id, nombre: 'Nueva partida', precioRef: 0 }]);
                }} className="px-3 py-1.5 rounded-lg text-xs font-semibold" style={{ backgroundColor: '#EEF1F4', color: C.slate }}>+ Agregar partida</button>
              </div>
            </div>

            <div style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 600, fontSize: 15, color: C.slate, marginTop: 24 }}>Tipos de servicio</div>
            <p style={{ fontSize: 12, color: C.muted, marginTop: 2 }}>Las opciones que aparecen al crear una orden de trabajo.</p>
            <div className="rounded-xl p-4 mt-3" style={{ backgroundColor: C.card, border: `1px solid ${C.border}`, maxWidth: 520 }}>
              <div className="flex flex-wrap gap-2">
                {tiposServicio.map((t) => (
                  <span key={t} className="flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg" style={{ backgroundColor: '#F6F8FA', fontSize: 12.5 }}>
                    {t}
                    {tiposServicio.length > 1 && <button onClick={() => { setTiposServicio((prev) => prev.filter((x) => x !== t)); supabase.from('tipos_servicio').delete().eq('empresa_id', empresaActiva).eq('nombre', t).then(() => {}); }}><X size={12} color={C.coral} /></button>}
                  </span>
                ))}
              </div>
              <div className="flex gap-2 items-end mt-3">
                <Field label="Nuevo tipo de servicio"><input style={{ ...inputStyle, width: 220 }} value={nuevoTipoServicio} onChange={(e) => setNuevoTipoServicio(e.target.value)} /></Field>
                <button onClick={async () => {
                  const v = nuevoTipoServicio.trim();
                  if (!v || tiposServicio.includes(v)) return;
                  const { error } = await supabase.from('tipos_servicio').insert({ empresa_id: empresaActiva, nombre: v });
                  if (error) { alert('No se pudo guardar: ' + error.message); return; }
                  setTiposServicio((prev) => [...prev, v]);
                  setNuevoTipoServicio('');
                }} className="px-3 py-2 rounded-lg text-xs font-semibold" style={{ backgroundColor: C.navy, color: '#fff' }}>Agregar</button>
              </div>
            </div>
            </>)}

            {configSeccion === 'cierre' && (
              <div className="rounded-xl p-5 mt-6" style={{ backgroundColor: C.card, border: `1px solid ${C.border}`, maxWidth: 640 }}>
                <div style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 600, fontSize: 15, color: C.slate }}>Cierre contable de periodo</div>
                <p style={{ fontSize: 12.5, color: C.muted, marginTop: 4 }}>Todos los comprobantes con fecha igual o anterior al cierre quedan protegidos: no se pueden crear, editar ni anular. Úsalo después de conciliar cada mes.</p>
                <div className="flex flex-wrap gap-3 items-end mt-4">
                  <Field label="Periodo cerrado hasta (inclusive)"><input type="date" style={inputStyle} value={cierreFecha} onChange={(e) => setCierreFecha(e.target.value)} /></Field>
                  {cierreFecha && <button onClick={() => pedirConfirmacion('¿Reabrir el periodo? Los comprobantes volverán a ser editables.', () => setCierreFecha(''))} className="px-3 py-2 rounded-lg text-sm font-semibold" style={{ backgroundColor: '#FDE8E8', color: C.coral }}>Reabrir periodo</button>}
                </div>
                {cierreFecha ? (
                  <div className="mt-4 rounded-xl p-3" style={{ backgroundColor: '#E4F8F1', border: `1px solid ${C.cyanDark}55`, fontSize: 12.5, color: '#0A6B5D' }}>🔒 La contabilidad está cerrada hasta el <strong>{cierreFecha}</strong>. Los asientos de esa fecha hacia atrás son de solo lectura.</div>
                ) : (
                  <div className="mt-4" style={{ fontSize: 12, color: C.muted }}>Sin cierre activo — todos los comprobantes son editables.</div>
                )}

                <div style={{ borderTop: `1px solid ${C.border}`, marginTop: 18, paddingTop: 16 }}>
                  <div style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 600, fontSize: 15, color: C.slate }}>Prorrateo de gastos administrativos del mes</div>
                  <p style={{ fontSize: 12.5, color: C.muted, marginTop: 4 }}>Distribuye los gastos del centro Administrativo entre los centros operativos y guarda constancia. La distribución también se ve en vivo en Informes → Estado de resultados por centro.</p>
                  <div className="flex flex-wrap gap-3 items-end mt-4">
                    <Field label="Mes"><input type="month" style={inputStyle} value={prorrateoMes} onChange={(e) => setProrrateoMes(e.target.value)} /></Field>
                    <Field label="Método">
                      <select style={{ ...inputStyle, width: 190 }} value={prorrateoMetodo} onChange={(e) => setProrrateoMetodo(e.target.value)}>
                        <option value="ventas">En función de ventas</option>
                        <option value="equitativo">Equitativo (partes iguales)</option>
                      </select>
                    </Field>
                    <button onClick={() => {
                      if (prorrateosGuardados.some((p) => p.periodo === prorrateoMes)) { alert(`El prorrateo de ${prorrateoMes} ya fue registrado.`); return; }
                      const desde = prorrateoMes + '-01';
                      const hasta = prorrateoMes + '-31';
                      const detalle = centros.filter((c) => c.codigo !== 'CC-ADM').map((c) => {
                        const T = erTotalsFor(desde, hasta, c.id);
                        return { centro: c.nombre, factor: T.factorProrrateo, monto: T.gastosAdmProrrateados };
                      });
                      const total = detalle.reduce((s, d) => s + d.monto, 0);
                      if (total === 0) { alert('No hay gastos administrativos en ese mes para prorratear.'); return; }
                      setProrrateosGuardados((prev) => [...prev, { periodo: prorrateoMes, metodo: prorrateoMetodo, detalle, total }]);
                    }} className="px-4 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.cyanDark }}>
                      Prorratear gastos del mes
                    </button>
                  </div>
                  {prorrateosGuardados.length > 0 && (
                    <div className="mt-4">
                      {[...prorrateosGuardados].reverse().map((p, i) => (
                        <div key={i} className="rounded-xl p-3 mb-2" style={{ backgroundColor: '#F6F8FA', border: `1px solid ${C.border}` }}>
                          <div style={{ fontSize: 12.5, fontWeight: 700, color: C.slate }}>{p.periodo} — método: {p.metodo === 'ventas' ? 'en función de ventas' : 'equitativo'} — total distribuido: <span style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(p.total)}</span></div>
                          {p.detalle.map((d, j) => (
                            <div key={j} className="flex justify-between" style={{ fontSize: 12, color: C.muted, marginTop: 2 }}>
                              <span>{d.centro} ({(d.factor * 100).toFixed(1)}%)</span>
                              <span style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(d.monto)}</span>
                            </div>
                          ))}
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              </div>
            )}

            <div style={{ fontSize: 11.5, color: C.muted, marginTop: 14 }}>Para dar de alta cuentas contables nuevas (para asociarlas aquí), ve a Contabilidad General → Plan de Cuentas.</div>
          </div>
        )}

        {/* AGENDA DE TECNICOS */}
        {tab === 'agenda' && (() => {
          const base = new Date(agendaBase + 'T00:00:00');
          const lunes = new Date(base); lunes.setDate(base.getDate() - ((base.getDay() + 6) % 7));
          const dias = Array.from({ length: 7 }, (_, i) => { const d = new Date(lunes); d.setDate(lunes.getDate() + i); return d; });
          const fstr = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
          const nombreDia = ['Lun', 'Mar', 'Mié', 'Jue', 'Vie', 'Sáb', 'Dom'];
          const mover = (delta) => { const d = new Date(lunes); d.setDate(lunes.getDate() + delta); setAgendaBase(fstr(d)); };
          const estadoColor = { pendiente: C.amber, en_proceso: '#3B82C4', completada: C.cyanDark, cancelada: C.coral };
          const cambiarVistaAgendaMovil = (v) => {
            setVistaAgendaMovil(v);
            if (sesion?.id) supabase.from('perfiles').update({ preferencias: { ...sesion.preferencias, vistaAgenda: v } }).eq('id', sesion.id).then(() => {});
          };
          const ordenesDia = (tecnicoId, fecha) => ordersVisiblesPorRol
            .filter((o) => o.tecnicoId === tecnicoId && o.fecha === fecha && o.estado !== 'por_validar')
            .filter((o) => agendaFiltro === 'todas' ? o.estado !== 'cancelada' : o.estado === agendaFiltro)
            .sort((a, b) => (a.hora || '99').localeCompare(b.hora || '99'));
          const hayChoque = (lista) => { const horas = lista.filter((o) => o.hora).map((o) => o.hora); return new Set(horas).size !== horas.length; };
          return (
            <div>
              <div className="flex items-center justify-between flex-wrap gap-3">
                <div>
                  <h1 style={{ fontFamily: 'Space Grotesk, sans-serif', fontSize: 22, fontWeight: 700, color: C.slate }}>{t('Agenda de técnicos')}</h1>
                  <p style={{ color: C.muted, fontSize: 13, marginTop: 2 }}>{t('Semana del')} {fstr(dias[0])} {t('al')} {fstr(dias[6])}.</p>
                </div>
                <div className="flex gap-2 items-center">
                  <button onClick={() => mover(-7)} className="px-3 py-2 rounded-lg text-sm font-semibold" style={{ backgroundColor: '#EEF1F4', color: C.slate }}>← {t('Semana anterior')}</button>
                  <button onClick={() => setAgendaBase(new Date().toISOString().slice(0, 10))} className="px-3 py-2 rounded-lg text-sm font-semibold" style={{ backgroundColor: C.navy, color: 'white' }}>{t('Hoy')}</button>
                  <button onClick={() => mover(7)} className="px-3 py-2 rounded-lg text-sm font-semibold" style={{ backgroundColor: '#EEF1F4', color: C.slate }}>{t('Semana siguiente')} →</button>
                </div>
              </div>

              <div className="flex gap-2 flex-wrap mt-4">
                {[
                  ['todas', t('Todas'), C.slate],
                  ['pendiente', t('Pendientes'), C.amber],
                  ['en_proceso', t('En proceso'), '#3B82C4'],
                  ['completada', t('Completadas'), C.cyanDark],
                  ['cancelada', t('Canceladas'), C.coral],
                ].map(([key, label, color]) => {
                  const count = key === 'todas'
                    ? orders.filter((o) => dias.some((d) => fstr(d) === o.fecha) && o.estado !== 'cancelada').length
                    : orders.filter((o) => dias.some((d) => fstr(d) === o.fecha) && o.estado === key).length;
                  const activo = agendaFiltro === key;
                  return (
                    <button key={key} onClick={() => setAgendaFiltro(key)}
                      className="px-3 py-1.5 rounded-full text-xs font-semibold"
                      style={{ backgroundColor: activo ? color : '#fff', color: activo ? '#fff' : color, border: `1.5px solid ${color}` }}>
                      {label} ({count})
                    </button>
                  );
                })}
              </div>

              <div className="hidden md:block rounded-xl overflow-hidden mt-4" style={{ backgroundColor: C.card, border: `1px solid ${C.border}`, overflowX: 'auto' }}>
                <table style={{ minWidth: 900 }}>
                  <thead>
                    <tr>
                      <th style={{ width: 130 }}>{t('Técnico')}</th>
                      {dias.map((d, i) => (
                        <th key={i} style={{ backgroundColor: fstr(d) === new Date().toISOString().slice(0, 10) ? '#E4F8F1' : undefined }}>
                          {nombreDia[i]} {d.getDate()}/{d.getMonth() + 1}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {techs.filter((t) => sesion.rol !== 'Tecnico' || !sesion.empleadoId || t.id === sesion.empleadoId).map((t) => (
                      <tr key={t.id}>
                        <td style={{ fontWeight: 600, fontSize: 12.5 }}>{t.nombre}<div style={{ fontSize: 10.5, color: C.muted, fontWeight: 400 }}>{orders.filter((o) => o.tecnicoId === t.id && dias.some((d) => fstr(d) === o.fecha) && o.estado !== 'cancelada').length} cita(s)</div></td>
                        {dias.map((d, i) => {
                          const lista = ordenesDia(t.id, fstr(d));
                          const choque = hayChoque(lista);
                          return (
                            <td key={i} style={{ verticalAlign: 'top', backgroundColor: choque ? '#FDE8E8' : undefined, minWidth: 110 }}>
                              {lista.map((o) => (
                                <div key={o.id} onClick={() => { setAgendaSel(o.id); setAgendaEdit(false); }} className="rounded px-1.5 py-1 mb-1" style={{ backgroundColor: '#F6F8FA', borderLeft: `3px solid ${estadoColor[o.estado]}`, fontSize: 10.5, cursor: 'pointer' }}>
                                  <div style={{ fontWeight: 700, fontFamily: 'JetBrains Mono, monospace' }}>{o.hora || 's/hora'} · {o.id.slice(-3)}</div>
                                  <div style={{ color: C.muted }}>{o.direccion?.slice(0, 28)}{o.direccion?.length > 28 ? '…' : ''}</div>
                                </div>
                              ))}
                              {choque && <div style={{ fontSize: 9.5, color: C.coral, fontWeight: 700 }}>⚠ Choque de horario</div>}
                            </td>
                          );
                        })}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <div className="flex gap-4 mt-3" style={{ fontSize: 11.5, color: C.muted }}>
                {Object.entries({ Pendiente: C.amber, 'En proceso': '#3B82C4', Completada: C.cyanDark }).map(([l, c]) => (
                  <span key={l} className="flex items-center gap-1.5"><span style={{ width: 10, height: 10, backgroundColor: c, borderRadius: 2, display: 'inline-block' }} />{l}</span>
                ))}
              </div>

              {/* Vistas móviles: cada quien elige cómo quiere ver su agenda en el celular, y el sistema lo recuerda. */}
              <div className="md:hidden mt-4">
                <div className="flex gap-2 mb-3">
                  <button onClick={() => cambiarVistaAgendaMovil('semana')} className="flex-1 py-2 rounded-lg text-sm font-semibold" style={{ backgroundColor: vistaAgendaMovil === 'semana' ? C.navy : '#fff', color: vistaAgendaMovil === 'semana' ? '#fff' : C.slate, border: `1px solid ${C.border}` }}>Semana</button>
                  <button onClick={() => cambiarVistaAgendaMovil('mes')} className="flex-1 py-2 rounded-lg text-sm font-semibold" style={{ backgroundColor: vistaAgendaMovil === 'mes' ? C.navy : '#fff', color: vistaAgendaMovil === 'mes' ? '#fff' : C.slate, border: `1px solid ${C.border}` }}>Mes</button>
                </div>

                {vistaAgendaMovil === 'semana' && (() => {
                  const propias = sesion.rol === 'Tecnico' && sesion.empleadoId ? ordersVisiblesPorRol.filter((o) => o.tecnicoId === sesion.empleadoId) : ordersVisiblesPorRol;
                  return (
                    <div className="flex flex-col gap-3">
                      {dias.map((d, i) => {
                        const fechaStr = fstr(d);
                        const citas = propias.filter((o) => o.fecha === fechaStr && o.estado !== 'cancelada' && o.estado !== 'por_validar').sort((a, b) => (a.hora || '99').localeCompare(b.hora || '99'));
                        const esHoy = fechaStr === new Date().toISOString().slice(0, 10);
                        return (
                          <div key={i} className="rounded-xl overflow-hidden" style={{ backgroundColor: C.card, border: `1px solid ${esHoy ? C.cyanDark : C.border}` }}>
                            <div style={{ backgroundColor: C.navy, color: '#fff', fontSize: 12, fontWeight: 600, padding: '7px 12px' }}>
                              {['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo'][i]} {d.getDate()}/{d.getMonth() + 1} · {citas.length} cita{citas.length === 1 ? '' : 's'}
                            </div>
                            {citas.length === 0 ? (
                              <div style={{ padding: '14px 12px', color: C.muted, fontSize: 12, textAlign: 'center' }}>Sin citas</div>
                            ) : citas.map((o, ci) => (
                              <div key={o.id} onClick={() => { setAgendaSel(o.id); setAgendaEdit(false); }} className="flex gap-2.5"
                                style={{ padding: '10px 12px', borderTop: ci > 0 ? `1px solid ${C.border}` : 'none', cursor: 'pointer' }}>
                                <div style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 12, color: C.muted, width: 40, flexShrink: 0 }}>{o.hora || '—'}</div>
                                <div style={{ flex: 1 }}>
                                  <div style={{ fontSize: 12.5, color: C.slate }}>{o.direccion || o.tipoServicio}</div>
                                  <div style={{ fontSize: 11, color: C.muted, marginTop: 1 }}>{o.tipoServicio}</div>
                                </div>
                                <div style={{ backgroundColor: estadoColor[o.estado] + '22', color: estadoColor[o.estado], fontSize: 10, fontWeight: 600, padding: '3px 7px', borderRadius: 6, height: 'fit-content' }}>{ESTADOS[o.estado]?.label}</div>
                              </div>
                            ))}
                          </div>
                        );
                      })}
                    </div>
                  );
                })()}

                {vistaAgendaMovil === 'mes' && (() => {
                  const propias = sesion.rol === 'Tecnico' && sesion.empleadoId ? ordersVisiblesPorRol.filter((o) => o.tecnicoId === sesion.empleadoId) : ordersVisiblesPorRol;
                  const base = new Date(mesAgendaBase + 'T00:00:00');
                  const anio = base.getFullYear(), mesNum = base.getMonth();
                  const primerDia = new Date(anio, mesNum, 1);
                  const diasEnMes = new Date(anio, mesNum + 1, 0).getDate();
                  const offset = (primerDia.getDay() + 6) % 7; // lunes = 0
                  const celdas = [...Array(offset).fill(null), ...Array.from({ length: diasEnMes }, (_, i) => i + 1)];
                  const fstrMes = (dia) => `${anio}-${String(mesNum + 1).padStart(2, '0')}-${String(dia).padStart(2, '0')}`;
                  const estadoDelDia = (dia) => {
                    const citas = propias.filter((o) => o.fecha === fstrMes(dia) && o.estado !== 'cancelada' && o.estado !== 'por_validar');
                    if (citas.length === 0) return null;
                    return citas.some((o) => o.estado === 'pendiente' || o.estado === 'en_proceso') ? C.amber : C.cyanDark;
                  };
                  const citasDelDiaSel = propias.filter((o) => o.fecha === diaMesSel && o.estado !== 'cancelada' && o.estado !== 'por_validar').sort((a, b) => (a.hora || '99').localeCompare(b.hora || '99'));
                  const moverMes = (delta) => { const d = new Date(anio, mesNum + delta, 1); setMesAgendaBase(d.toISOString().slice(0, 10)); };
                  return (
                    <div>
                      <div className="flex items-center justify-between mb-2">
                        <button onClick={() => moverMes(-1)} style={{ color: C.cyanDark, fontSize: 15, padding: '4px 8px' }}>←</button>
                        <div style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 700, fontSize: 14, color: C.slate }}>
                          {['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'][mesNum]} {anio}
                        </div>
                        <button onClick={() => moverMes(1)} style={{ color: C.cyanDark, fontSize: 15, padding: '4px 8px' }}>→</button>
                      </div>
                      <div className="rounded-xl p-2.5" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
                        <div className="grid grid-cols-7 keep-grid-cols" style={{ textAlign: 'center', marginBottom: 4 }}>
                          {['L', 'M', 'X', 'J', 'V', 'S', 'D'].map((d) => <div key={d} style={{ fontSize: 10, color: '#9AA5B1', fontWeight: 600 }}>{d}</div>)}
                        </div>
                        <div className="grid grid-cols-7 keep-grid-cols" style={{ rowGap: 4 }}>
                          {celdas.map((dia, idx) => {
                            if (!dia) return <div key={idx} />;
                            const fechaStr = fstrMes(dia);
                            const color = estadoDelDia(dia);
                            const seleccionado = fechaStr === diaMesSel;
                            return (
                              <div key={idx} onClick={() => setDiaMesSel(fechaStr)} style={{ textAlign: 'center', padding: '6px 0', borderRadius: 8, cursor: 'pointer', backgroundColor: seleccionado ? C.cyanDark : 'transparent', color: seleccionado ? '#fff' : C.slate, fontSize: 12 }}>
                                {dia}
                                {!seleccionado && color && <div style={{ width: 4, height: 4, borderRadius: '50%', backgroundColor: color, margin: '2px auto 0' }} />}
                              </div>
                            );
                          })}
                        </div>
                      </div>
                      <div className="flex gap-3 mt-2.5" style={{ fontSize: 10.5, color: C.muted }}>
                        <span><span style={{ display: 'inline-block', width: 7, height: 7, borderRadius: '50%', backgroundColor: C.amber, marginRight: 4 }} />Pendiente</span>
                        <span><span style={{ display: 'inline-block', width: 7, height: 7, borderRadius: '50%', backgroundColor: C.cyanDark, marginRight: 4 }} />Completada</span>
                      </div>

                      <div className="rounded-xl overflow-hidden mt-3" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
                        <div style={{ backgroundColor: C.navy, color: '#fff', fontSize: 12, fontWeight: 600, padding: '7px 12px' }}>
                          {diaMesSel} · {citasDelDiaSel.length} cita{citasDelDiaSel.length === 1 ? '' : 's'}
                        </div>
                        {citasDelDiaSel.length === 0 ? (
                          <div style={{ padding: '14px 12px', color: C.muted, fontSize: 12, textAlign: 'center' }}>Sin citas</div>
                        ) : citasDelDiaSel.map((o, ci) => (
                          <div key={o.id} onClick={() => { setAgendaSel(o.id); setAgendaEdit(false); }} className="flex gap-2.5"
                            style={{ padding: '10px 12px', borderTop: ci > 0 ? `1px solid ${C.border}` : 'none', cursor: 'pointer' }}>
                            <div style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 12, color: C.muted, width: 40, flexShrink: 0 }}>{o.hora || '—'}</div>
                            <div style={{ flex: 1 }}>
                              <div style={{ fontSize: 12.5, color: C.slate }}>{o.direccion || o.tipoServicio}</div>
                              <div style={{ fontSize: 11, color: C.muted, marginTop: 1 }}>{o.tipoServicio}</div>
                            </div>
                            <div style={{ backgroundColor: estadoColor[o.estado] + '22', color: estadoColor[o.estado], fontSize: 10, fontWeight: 600, padding: '3px 7px', borderRadius: 6, height: 'fit-content' }}>{ESTADOS[o.estado]?.label}</div>
                          </div>
                        ))}
                      </div>
                    </div>
                  );
                })()}
              </div>

              {agendaSel && (() => {
                const o = orders.find((x) => x.id === agendaSel);
                if (!o) return null;
                const iniciarEdicion = () => { setAgendaDraft({ tecnicoId: o.tecnicoId, fecha: o.fecha, hora: o.hora || '', direccion: o.direccion || '', notas: o.notas || '' }); setAgendaEdit(true); };
                const guardar = () => {
                  const cambioFechaHora = agendaDraft.fecha !== o.fecha || (agendaDraft.hora || '') !== (o.hora || '');
                  const nuevasReprogramaciones = cambioFechaHora
                    ? [...(o.reprogramaciones || []), { de: `${o.fecha}${o.hora ? ' ' + o.hora : ''}`, a: `${agendaDraft.fecha}${agendaDraft.hora ? ' ' + agendaDraft.hora : ''}`, registrado: new Date().toISOString().slice(0, 16).replace('T', ' '), por: sesion.nombre }]
                    : o.reprogramaciones;
                  setOrders((prev) => prev.map((x) => x.id === o.id ? {
                    ...x, tecnicoId: agendaDraft.tecnicoId, fecha: agendaDraft.fecha, hora: agendaDraft.hora || null, direccion: agendaDraft.direccion, notas: agendaDraft.notas,
                    reprogramaciones: nuevasReprogramaciones,
                  } : x));
                  setAgendaEdit(false);
                  if (o.supabaseId) {
                    supabase.from('ordenes').update({
                      tecnico_id: agendaDraft.tecnicoId || null, fecha: agendaDraft.fecha, hora: agendaDraft.hora || null,
                      direccion: agendaDraft.direccion, notas: agendaDraft.notas, reprogramaciones: nuevasReprogramaciones,
                    }).eq('id', o.supabaseId).then(({ error }) => { if (error) console.error('Error al guardar cambios de agenda:', error.message); });
                  }
                };
                const estadoLabel = { pendiente: 'Pendiente', en_proceso: 'En proceso', completada: 'Completada', cancelada: 'Cancelada' };
                const Row = ({ label, children }) => (
                  <div className="flex justify-between items-center py-1.5" style={{ borderBottom: `1px solid ${C.border}` }}>
                    <span style={{ fontSize: 12, color: C.muted, fontWeight: 600 }}>{label}</span>
                    <span style={{ fontSize: 12.5, color: C.slate, textAlign: 'right', maxWidth: '60%' }}>{children}</span>
                  </div>
                );
                return (
                  <div style={{ position: 'fixed', inset: 0, backgroundColor: 'rgba(11,22,34,0.45)', zIndex: 50, display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 16 }}
                    onClick={() => { setAgendaSel(null); setAgendaEdit(false); }}>
                    <div onClick={(e) => e.stopPropagation()} className="rounded-2xl p-5" style={{ backgroundColor: '#fff', width: 440, maxWidth: '100%', maxHeight: '85vh', overflowY: 'auto', boxShadow: '0 20px 50px rgba(0,0,0,0.25)' }}>
                      <div className="flex items-center justify-between mb-3">
                        <div>
                          <div style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 700, fontSize: 16, color: C.slate }}>{o.id}</div>
                          {o.referencia && <div style={{ fontSize: 11, color: C.muted }}>{o.referencia}</div>}
                        </div>
                        <div className="flex items-center gap-2">
                          <span className="text-xs font-semibold px-2 py-0.5 rounded-full" style={{ backgroundColor: '#F6F8FA', color: estadoColor[o.estado], border: `1.5px solid ${estadoColor[o.estado]}` }}>{estadoLabel[o.estado]}</span>
                          {!agendaEdit && <button onClick={iniciarEdicion} title="Modificar orden" className="p-1.5 rounded-lg" style={{ backgroundColor: '#EEF1F4' }}><Pencil size={15} color={C.slate} /></button>}
                          <button onClick={() => { setAgendaSel(null); setAgendaEdit(false); }} className="p-1.5 rounded-lg" style={{ backgroundColor: '#EEF1F4' }}><X size={15} color={C.slate} /></button>
                        </div>
                      </div>

                      {!agendaEdit ? (
                        <div>
                          <Row label={t('Cliente')}>{clientName(o.clienteId)}</Row>
                          <Row label={t('Técnico')}>{techName(o.tecnicoId)}</Row>
                          <Row label="Estado">
                            {o.estado === 'por_validar' ? (
                              <span style={{ fontWeight: 600, color: C.amber }}>Pendiente de validación</span>
                            ) : esSoloLectura(sesion.rol, 'agenda') ? (
                              <span style={{ fontWeight: 600 }}>{ESTADOS[o.estado]?.label || o.estado}</span>
                            ) : (
                              <select value={o.estado} onChange={(e) => changeOrderEstado(o.id, e.target.value)} style={{ ...inputStyle, fontSize: 12.5, padding: '4px 8px' }}>
                                {['pendiente', 'en_proceso', 'completada', 'cancelada'].map((es) => (
                                  <option key={es} value={es}>{ESTADOS[es]?.label || es}</option>
                                ))}
                              </select>
                            )}
                          </Row>
                          <Row label={t('Tipo de servicio')}>{o.tipoServicio}</Row>
                          <Row label={t('Dirección')}>{o.direccion || '—'}</Row>
                          <Row label={t('Fecha y hora')}>{o.fecha}{o.hora ? ` · ${o.hora}` : ''}</Row>
                          <Row label={t('Centro de costo')}>{centroCostoName(o.centroCostoId)}</Row>
                          {!ocultaValores('agenda') && <Row label={t('Precio al cliente')}>{fmt(o.precioCliente)} <span style={{ color: C.muted, fontSize: 10.5 }}>({o.estadoPago === 'pagada' ? t('cobrado') : t('por cobrar')})</span></Row>}
                          {!ocultaValores('agenda') && <Row label={t('Mano de obra')}>{fmt(o.costoManoObra)} <span style={{ color: C.muted, fontSize: 10.5 }}>({o.estadoPagoTecnico === 'pagado' ? t('pagada') : t('por pagar')})</span></Row>}
                          {o.squadra && <Row label="Squadra (equipo OLO)">{o.squadra}</Row>}
                          {o.idRisorsa && <Row label="ID Risorsa">{o.idRisorsa}</Row>}
                          {(o.building || o.cnoBuilding || o.roeBuilding) && (
                            <Row label="Building / CNO / ROE">{[o.building, o.cnoBuilding, o.roeBuilding].filter(Boolean).join(' · ')}</Row>
                          )}
                          {o.notas && <Row label={t('Notas')}>{o.notas}</Row>}
                          {(o.reprogramaciones || []).length > 0 && (
                            <div style={{ marginTop: 10 }}>
                              <div style={{ fontSize: 11, fontWeight: 700, color: C.muted, textTransform: 'uppercase' }}>{t('Historial de reprogramaciones')}</div>
                              {o.reprogramaciones.map((rp, i) => (
                                <div key={i} style={{ fontSize: 11.5, color: C.slate, marginTop: 3, backgroundColor: '#FFF7E8', borderRadius: 6, padding: '4px 8px' }}>
                                  {rp.de} → <strong>{rp.a}</strong> <span style={{ color: C.muted }}>· {rp.registrado} · {rp.por}</span>
                                </div>
                              ))}
                            </div>
                          )}
                        </div>
                      ) : (
                        <div>
                          {sesion.rol !== 'Tecnico' && (
                            <>
                              <Field label={t('Técnico')}>
                                <select style={{ ...inputStyle, width: '100%' }} value={agendaDraft.tecnicoId} onChange={(e) => setAgendaDraft({ ...agendaDraft, tecnicoId: e.target.value })}>
                                  {techs.map((t) => <option key={t.id} value={t.id}>{t.nombre}</option>)}
                                </select>
                              </Field>
                              <div className="flex gap-3 mt-3">
                                <Field label={t('Fecha')}><input type="date" style={inputStyle} value={agendaDraft.fecha} onChange={(e) => setAgendaDraft({ ...agendaDraft, fecha: e.target.value })} /></Field>
                                <Field label="Hora"><input type="time" style={inputStyle} value={agendaDraft.hora} onChange={(e) => setAgendaDraft({ ...agendaDraft, hora: e.target.value })} /></Field>
                              </div>
                              <div className="mt-3"><Field label={t('Dirección')}><input style={{ ...inputStyle, width: '100%' }} value={agendaDraft.direccion} onChange={(e) => setAgendaDraft({ ...agendaDraft, direccion: e.target.value })} /></Field></div>
                            </>
                          )}
                          <div className="mt-3"><Field label={t('Notas')}><input style={{ ...inputStyle, width: '100%' }} value={agendaDraft.notas} onChange={(e) => setAgendaDraft({ ...agendaDraft, notas: e.target.value })} /></Field></div>
                          {sesion.rol !== 'Tecnico' && <div style={{ fontSize: 11, color: C.muted, marginTop: 10 }}>{t('Los montos y el estado se gestionan desde la pestaña Órdenes.')}</div>}
                          <div className="flex gap-2 mt-4">
                            <button onClick={guardar} className="px-4 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.cyanDark }}>{t('Guardar cambios')}</button>
                            <button onClick={() => setAgendaEdit(false)} className="px-4 py-2 rounded-lg text-sm font-semibold" style={{ backgroundColor: '#EEF1F4', color: C.slate }}>{t('Cancelar')}</button>
                          </div>
                        </div>
                      )}
                    </div>
                  </div>
                );
              })()}
            </div>
          );
        })()}

        {/* VENTAS (giro Comercio) */}
        {tab === 'ventas' && (() => {
          const vigentes = ventas.filter((v) => !v.anulada);
          const porCobrar = vigentes.filter((v) => v.estadoPago === 'pendiente').reduce((s, v) => s + v.total, 0);
          const mesActual = new Date().toISOString().slice(0, 7);
          const ventasMes = vigentes.filter((v) => v.fecha.startsWith(mesActual)).reduce((s, v) => s + v.total, 0);
          const margenMes = vigentes.filter((v) => v.fecha.startsWith(mesActual)).reduce((s, v) => s + (v.total - v.costo), 0);
          return (
          <div>
            <div className="flex items-center justify-between">
              <div>
                <h1 style={{ fontFamily: 'Space Grotesk, sans-serif', fontSize: 22, fontWeight: 700, color: C.slate }}>{t('Ventas')}</h1>
                <p style={{ color: C.muted, fontSize: 13, marginTop: 2 }}>Cada venta descarga el inventario al costo promedio y contabiliza ingreso y costo de ventas automáticamente.</p>
              </div>
              <button onClick={() => setShowVentaForm(!showVentaForm)} className="flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.navy }}>
                {showVentaForm ? <X size={15} /> : <Plus size={15} />} {showVentaForm ? 'Cerrar' : 'Nueva venta'}
              </button>
            </div>

            {showVentaForm && <VentaForm clients={clients.filter((c) => c.activo !== false)} materials={materials} onSubmit={(data) => { addVenta(data); setShowVentaForm(false); }} />}

            <div className="grid grid-cols-3 gap-4 mt-5">
              {[
                ['Ventas del mes', fmt(ventasMes), C.cyanDark],
                ['Margen bruto del mes', fmt(margenMes), margenMes >= 0 ? C.cyanDark : C.coral],
                ['Por cobrar', fmt(porCobrar), porCobrar > 0 ? C.amber : C.cyanDark],
              ].map(([label, val, color], i) => (
                <div key={i} className="rounded-xl p-4" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
                  <div style={{ fontSize: 11, fontWeight: 700, color: C.muted, textTransform: 'uppercase' }}>{label}</div>
                  <div style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 20, fontWeight: 700, color, marginTop: 6 }}>{val}</div>
                </div>
              ))}
            </div>

            <div className="rounded-xl overflow-hidden mt-5" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
              <table>
                <thead><tr><th>Folio</th><th>Cliente</th><th>Fecha</th><th>Productos</th><th>Total</th><th>Margen</th><th>Cobro</th><th></th></tr></thead>
                <tbody>
                  {ventasPag.paginaActual.map((v) => (
                    <tr key={v.id} style={{ opacity: v.anulada ? 0.5 : 1 }}>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 12 }}>{v.id}{v.anulada && <span style={{ color: C.coral, fontSize: 10, display: 'block' }}>ANULADA</span>}</td>
                      <td>{clientName(v.clienteId)}</td>
                      <td>{v.fecha}</td>
                      <td style={{ fontSize: 11.5 }}>{v.lineas.map((l) => `${l.cantidad}× ${matById(l.materialId)?.nombre || l.materialId}`).join(', ')}</td>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(v.total)}</td>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 12, color: v.total - v.costo >= 0 ? C.cyanDark : C.coral }}>{fmt(v.total - v.costo)}</td>
                      <td>{!v.anulada && <PayPill ok={v.estadoPago === 'pagada'} textOk="Cobrada" textPending="Marcar cobrada" onClick={() => toggleVentaCobro(v.id)} />}</td>
                      <td>{!v.anulada && v.estadoPago !== 'pagada' && <button onClick={() => anularVenta(v.id)} title="Anular venta"><X size={14} color={C.coral} /></button>}</td>
                    </tr>
                  ))}
                  {ventas.length === 0 && <tr><td colSpan={8} style={{ color: C.muted, textAlign: 'center', padding: 24 }}>Sin ventas registradas. Da de alta productos en Materiales y registra tu primera venta.</td></tr>}
                </tbody>
                {ventas.filter((v) => !v.anulada).length > 0 && (
                  <tfoot>
                    <tr>
                      <td colSpan={4} style={{ fontWeight: 700 }}>Total ({ventas.filter((v) => !v.anulada).length} vigente{ventas.filter((v) => !v.anulada).length === 1 ? '' : 's'})</td>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace', fontWeight: 700 }}>{fmt(ventas.filter((v) => !v.anulada).reduce((s, v) => s + v.total, 0))}</td>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace', fontWeight: 700 }}>{fmt(ventas.filter((v) => !v.anulada).reduce((s, v) => s + (v.total - v.costo), 0))}</td>
                      <td></td><td></td>
                    </tr>
                  </tfoot>
                )}
              </table>
              <Paginador p={ventasPag} />
            </div>
          </div>
          );
        })()}

        {/* EMPRESAS (multiempresa) */}
        {tab === 'empresas' && sesion.rol === 'Administrador' && (
          <div>
            <div className="flex items-center justify-between">
              <div>
                <h1 style={{ fontFamily: 'Space Grotesk, sans-serif', fontSize: 22, fontWeight: 700, color: C.slate }}>{t('Empresas')}</h1>
                <p style={{ color: C.muted, fontSize: 13, marginTop: 2 }}>Cada empresa tiene su propio espacio completo: contabilidad, órdenes, catálogos y cierres independientes. Puedes registrar empresas de cualquier giro.</p>
              </div>
              <div className="flex gap-2 flex-wrap">
                <button onClick={descargarRespaldoCompleto} className="flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-sm font-semibold" style={{ backgroundColor: '#fff', color: C.slate, border: `1px solid ${C.border}` }}>
                  <Download size={15} /> Respaldo completo
                </button>
                <label className="flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-sm font-semibold cursor-pointer" style={{ backgroundColor: '#fff', color: C.slate, border: `1px solid ${C.border}` }}>
                  <Download size={15} style={{ transform: 'rotate(180deg)' }} /> Restaurar respaldo
                  <input type="file" accept=".json" style={{ display: 'none' }} onChange={(e) => { if (e.target.files[0]) restaurarRespaldo(e.target.files[0]); e.target.value = ''; }} />
                </label>
                <button onClick={iniciarDeCero} className="flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-sm font-semibold" style={{ backgroundColor: '#FDE8E8', color: C.coral, border: `1px solid ${C.coral}55` }}>
                  <Trash2 size={15} /> Iniciar de cero
                </button>
                <button onClick={() => setShowEmpresaForm(!showEmpresaForm)} className="flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.navy }}>
                  {showEmpresaForm ? <X size={15} /> : <Plus size={15} />} {showEmpresaForm ? 'Cerrar' : 'Nueva empresa'}
                </button>
              </div>
            </div>

            <div className="rounded-xl p-3 mt-4" style={{ backgroundColor: '#F0FBF9', border: `1px solid ${C.cyanDark}33` }}>
              <div style={{ fontSize: 12, color: C.slate }}>
                💾 <strong>Respaldo completo</strong> descarga un solo archivo JSON con absolutamente todo. <strong>Restaurar respaldo</strong> vuelve a subir ese mismo archivo y reemplaza toda la información actual — así es como subes tu información inicial real: exporta un respaldo, edítalo o constrúyelo con tus datos, y restáuralo aquí.
              </div>
              <div style={{ fontSize: 11, color: C.muted, marginTop: 4 }}>⚠ Ambas acciones manejan todas las claves de acceso en texto plano (en memoria, sin encriptar) — guarda estos archivos con el mismo cuidado que una contraseña.</div>
            </div>
            <div className="rounded-xl p-3 mt-3" style={{ backgroundColor: '#FDE8E8', border: `1px solid ${C.coral}33` }}>
              <div style={{ fontSize: 12, color: C.slate }}>
                🗑 <strong>Iniciar de cero</strong> vacía por completo la empresa en la que estás ahora mismo ({empresaActual?.nombre}) — borra sus órdenes, contabilidad, clientes, empleados e inventario de ejemplo, y la deja lista para capturar información real desde cero. No afecta a otras empresas.
              </div>
            </div>

            {showEmpresaForm && <EmpresaForm onSubmit={(data) => { addEmpresa(data); setShowEmpresaForm(false); }} />}

            <div className="grid grid-cols-2 gap-3 mt-5">
              {empresas.map((e) => (
                <div key={e.id} className="rounded-xl p-4" style={{ backgroundColor: C.card, border: e.id === empresaActiva ? `2px solid ${C.cyanDark}` : `1px solid ${C.border}`, opacity: e.activa === false ? 0.55 : 1 }}>
                  <div className="flex items-center justify-between">
                    <div>
                      <div className="flex items-center gap-2">
                        <span style={{ fontWeight: 700, fontSize: 15, color: C.slate }}>{e.nombre}</span>
                        {e.id === empresaActiva && <span className="text-xs font-bold px-2 py-0.5 rounded-full" style={{ backgroundColor: '#E4F8F1', color: C.cyanDark }}>ACTIVA</span>}
                        {e.activa === false && <span className="text-xs font-bold px-2 py-0.5 rounded-full" style={{ backgroundColor: '#FDE8E8', color: C.coral }}>SUSPENDIDA</span>}
                      </div>
                      <div style={{ fontSize: 12, color: C.muted, marginTop: 3 }}>🏢 {e.giro}</div>
                      {e.eslogan && <div style={{ fontSize: 11.5, color: C.muted, fontStyle: 'italic', marginTop: 1 }}>"{e.eslogan}"</div>}
                    </div>
                    <div className="flex flex-col gap-1.5 items-end">
                      {e.id !== empresaActiva && e.activa !== false && (
                        <button onClick={() => cambiarEmpresa(e.id)} className="text-xs font-semibold px-3 py-1.5 rounded-lg text-white" style={{ backgroundColor: C.cyanDark }}>Entrar</button>
                      )}
                      {e.id !== empresaActiva && (
                        <button onClick={() => setEmpresas((prev) => prev.map((x) => x.id === e.id ? { ...x, activa: x.activa === false } : x))} className="text-xs font-semibold px-2 py-1 rounded-lg" style={{ backgroundColor: e.activa === false ? '#E4F8F1' : '#FDE8E8', color: e.activa === false ? C.cyanDark : C.coral }}>
                          {e.activa === false ? 'Reactivar' : 'Suspender'}
                        </button>
                      )}
                    </div>
                  </div>
                  <div style={{ fontSize: 11, color: C.muted, marginTop: 8 }}>
                    {e.id === empresaActiva
                      ? `Datos en uso: ${orders.length} órdenes · ${asientos.length} asientos · ${empleados.length} empleados`
                      : 'Datos guardados en Supabase — entra para verlos'}
                  </div>
                </div>
              ))}
            </div>
            <div style={{ fontSize: 11.5, color: C.muted, marginTop: 12 }}>Las empresas nuevas arrancan con un plan de cuentas y categorías genéricas que funcionan para cualquier giro — cada una las adapta en su propia Configuración. Los usuarios y roles son compartidos; el acceso por empresa se asigna en Usuarios y roles.</div>
          </div>
        )}

        {/* USUARIOS Y ROLES */}
        {tab === 'usuarios' && sesion.rol === 'Administrador' && (
          <div>
            <div>
              <h1 style={{ fontFamily: 'Space Grotesk, sans-serif', fontSize: 22, fontWeight: 700, color: C.slate }}>{t('Usuarios y roles')}</h1>
              <p style={{ color: C.muted, fontSize: 13, marginTop: 2 }}>Registra el correo de quien necesite acceso a esta empresa. Aún no se envía un correo automático — avísale tú que entre al sitio y se registre (o inicie sesión) con ese mismo correo; ahí se le da acceso solo.</p>
            </div>

            <div className="rounded-xl p-4 mt-4 flex flex-wrap gap-3 items-end" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
              <Field label="Correo a invitar"><input type="email" style={{ ...inputStyle, width: 240 }} value={inviteEmail} onChange={(e) => setInviteEmail(e.target.value)} placeholder="persona@correo.com" /></Field>
              <Field label="Rol">
                <select style={{ ...inputStyle, width: 180 }} value={inviteRol} onChange={(e) => setInviteRol(e.target.value)}>
                  <option value="">Elegir…</option>
                  {rolesLista.filter((r) => r !== 'Administrador' || sesion.rol === 'Administrador').map((r) => <option key={r} value={r}>{r}</option>)}
                </select>
              </Field>
              <button onClick={invitarUsuario} className="px-4 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.cyanDark }}>Enviar invitación</button>
            </div>

            <div className="rounded-xl overflow-hidden mt-5" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
              <div className="px-4 pt-4 pb-1" style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 600, fontSize: 14, color: C.slate }}>Personas con acceso</div>
              <table>
                <thead><tr><th>Nombre</th><th>Correo</th><th>Rol</th><th>Técnico vinculado</th><th>Ve avisos de bancos</th><th>Último acceso</th><th></th></tr></thead>
                <tbody>
                  {miembrosEmpresa.map((m) => (
                    <tr key={m.perfilId}>
                      <td style={{ fontWeight: 600 }}>
                        <input defaultValue={m.nombre || ''} onBlur={(e) => { if (e.target.value.trim() && e.target.value.trim() !== m.nombre) cambiarNombreMiembro(m.perfilId, e.target.value.trim()); }}
                          style={{ ...inputStyle, fontSize: 12.5, padding: '4px 6px', width: 140, fontWeight: 600, border: 'none', backgroundColor: 'transparent' }} />
                        {m.perfilId === sesion.id && <span style={{ fontSize: 10.5, color: C.cyanDark, marginLeft: 2 }}>(tú)</span>}
                      </td>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 12 }}>{m.email}</td>
                      <td>
                        <select value={m.rol || ''} onChange={(e) => cambiarRolMiembro(m.perfilId, e.target.value)} disabled={m.perfilId === sesion.id} style={{ ...inputStyle, fontSize: 12, padding: '4px 6px' }}>
                          {rolesLista.map((r) => <option key={r} value={r}>{r}</option>)}
                        </select>
                      </td>
                      <td>
                        {m.rol === 'Tecnico' ? (
                          <select value={m.empleadoId || ''} onChange={(e) => vincularEmpleadoMiembro(m.perfilId, e.target.value || null)} style={{ ...inputStyle, fontSize: 12, padding: '4px 6px' }}>
                            <option value="">— Sin vincular —</option>
                            {techs.map((t) => <option key={t.id} value={t.id}>{t.nombre}</option>)}
                          </select>
                        ) : <span style={{ color: C.muted, fontSize: 12 }}>—</span>}
                      </td>
                      <td style={{ textAlign: 'center' }}>
                        <input type="checkbox" checked={!!m.recibeAvisosFinancieros} onChange={(e) => toggleAvisosFinancieros(m.perfilId, e.target.checked)} style={{ width: 15, height: 15, accentColor: C.cyanDark }} />
                      </td>
                      <td style={{ fontSize: 11.5, color: C.muted }}>{m.ultimoAcceso ? new Date(m.ultimoAcceso).toLocaleString() : 'Nunca'}</td>
                      <td>{m.perfilId !== sesion.id && <button onClick={() => quitarAccesoMiembro(m.perfilId)} title="Quitar acceso"><Trash2 size={14} color={C.coral} /></button>}</td>
                    </tr>
                  ))}
                  {miembrosEmpresa.length === 0 && <tr><td colSpan={6} style={{ color: C.muted, textAlign: 'center', padding: 20 }}>Sin miembros todavía.</td></tr>}
                </tbody>
              </table>
            </div>

            {invitacionesEmpresa.length > 0 && (
              <div className="rounded-xl overflow-hidden mt-5" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
                <div className="px-4 pt-4 pb-1" style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 600, fontSize: 14, color: C.slate }}>Invitaciones pendientes</div>
                <table>
                  <thead><tr><th>Correo</th><th>Rol</th><th></th></tr></thead>
                  <tbody>
                    {invitacionesEmpresa.map((i) => (
                      <tr key={i.id}>
                        <td style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 12 }}>{i.email}</td>
                        <td style={{ fontSize: 12 }}>{i.rol}</td>
                        <td><button onClick={() => cancelarInvitacion(i.id)} title="Cancelar invitación"><X size={14} color={C.coral} /></button></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}

            <div className="rounded-xl p-4 mt-5" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
              <div className="flex items-center justify-between flex-wrap gap-3">
                <div>
                  <div style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 600, fontSize: 14, color: C.slate }}>Matriz de permisos por rol</div>
                  <div style={{ fontSize: 12, color: C.muted, marginTop: 2 }}>Marca qué módulos ve cada rol. Los cambios aplican de inmediato. El Administrador siempre tiene acceso total.</div>
                </div>
                <div className="flex gap-2 items-end">
                  <Field label="Nuevo rol"><input style={{ ...inputStyle, width: 150 }} value={nuevoRolNombre} onChange={(e) => setNuevoRolNombre(e.target.value)} placeholder="Ej. Supervisor" /></Field>
                  <button onClick={() => {
                    const nombre = nuevoRolNombre.trim();
                    if (!nombre) return;
                    addRol(nombre);
                    setNuevoRolNombre('');
                  }} className="px-3 py-2 rounded-lg text-white text-xs font-semibold" style={{ backgroundColor: C.navy }}>Crear rol</button>
                </div>
              </div>

              <div style={{ overflowX: 'auto', marginTop: 12 }}>
                <table style={{ minWidth: 640 }}>
                  <thead>
                    <tr>
                      <th>Módulo</th>
                      {rolesLista.map((r) => (
                        <th key={r} style={{ textAlign: 'center' }}>
                          {r}
                          {r !== 'Administrador' && !['Contabilidad', 'Operaciones', 'Tecnico'].includes(r) && !miembrosEmpresa.some((m) => m.rol === r) && (
                            <button onClick={() => pedirConfirmacion(`¿Eliminar el rol ${r}?`, () => eliminarRol(r))} title="Eliminar rol" style={{ marginLeft: 4 }}><Trash2 size={11} color={C.coral} /></button>
                          )}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {MODULOS_PERMISOS.map((g) => (
                      <React.Fragment key={g.grupo}>
                        <tr><td colSpan={1 + rolesLista.length} style={{ backgroundColor: '#F6F8FA', fontSize: 11, fontWeight: 700, color: C.muted, textTransform: 'uppercase', padding: '6px 12px' }}>{g.grupo}</td></tr>
                        {g.items.map(([key, label]) => (
                          <tr key={key}>
                            <td style={{ fontSize: 12.5 }}>{label}</td>
                            {rolesLista.map((r) => (
                              <td key={r} style={{ textAlign: 'center' }}>
                                {r === 'Administrador'
                                  ? <span style={{ color: C.cyanDark, fontWeight: 700 }}>✓</span>
                                  : (
                                    <div className="flex flex-col items-center gap-1">
                                      <select
                                        value={(rolesConfig[r] || []).includes(key) ? 'editar' : (rolesLectura[r] || []).includes(key) ? 'ver' : 'no'}
                                        onChange={(e) => setPermisoNivel(r, key, e.target.value)}
                                        style={{ ...inputStyle, fontSize: 11, padding: '2px 4px', width: 78, color: (rolesConfig[r] || []).includes(key) ? C.cyanDark : (rolesLectura[r] || []).includes(key) ? '#3B82C4' : C.muted }}>
                                        <option value="no">— No —</option>
                                        <option value="ver">Ver</option>
                                        <option value="editar">Editar</option>
                                      </select>
                                      {(tieneAcceso(r, key)) && (
                                        <label className="flex items-center gap-1" title="Si está marcado, este rol ve los montos en $ de este módulo" style={{ fontSize: 9.5, color: C.muted }}>
                                          <input type="checkbox" checked={!(rolesOcultarValores[r] || []).includes(key)} onChange={() => toggleOcultarValores(r, key)} style={{ width: 11, height: 11, accentColor: C.cyanDark }} />
                                          ver $
                                        </label>
                                      )}
                                    </div>
                                  )}
                              </td>
                            ))}
                          </tr>
                        ))}
                      </React.Fragment>
                    ))}
                  </tbody>
                </table>
              </div>
              <div style={{ fontSize: 11.5, color: C.muted, marginTop: 10 }}>El rol "Tecnico" además restringe a ver solo las órdenes propias del técnico vinculado. Los roles personalizados solo se pueden eliminar si ningún usuario los tiene asignados.</div>
              <div style={{ fontSize: 11.5, color: C.amber, marginTop: 4 }}>⚠ Nota: en esta versión las claves y permisos se guardan en memoria del navegador. Al migrar a Supabase tendrán encriptación y autenticación real.</div>
            </div>
          </div>
        )}

        {/* CUENTA (personal, cualquier usuario) */}
        {tab === 'cuenta' && (
          <div>
            <h1 style={{ fontFamily: 'Space Grotesk, sans-serif', fontSize: 22, fontWeight: 700, color: C.slate }}>{t('Mi cuenta')}</h1>
            <p style={{ color: C.muted, fontSize: 13, marginTop: 2 }}>Administra tu propia cuenta — no necesitas al administrador para esto.</p>

            <div className="rounded-xl p-5 mt-5" style={{ backgroundColor: C.card, border: `1px solid ${C.border}`, maxWidth: 480 }}>
              <div className="flex items-center gap-3 mb-4">
                <div style={{ width: 46, height: 46, borderRadius: '50%', backgroundColor: C.navy, color: '#fff', display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 700 }}>
                  {sesion.nombre.split(' ').map((p) => p[0]).slice(0, 2).join('')}
                </div>
                <div>
                  <div style={{ fontWeight: 700, color: C.slate }}>@{sesion.username || sesion.usuario}</div>
                  <div style={{ fontSize: 12, color: C.cyanDark, fontWeight: 600 }}>{sesion.rol}</div>
                </div>
              </div>

              <Field label="Nombre completo"><input style={{ ...inputStyle, width: '100%' }} value={cuentaNombreEdit} onChange={(e) => setCuentaNombreEdit(e.target.value)} /></Field>
              <div className="mt-3">
                <Field label="Usuario (para iniciar sesión sin escribir tu correo)">
                  <input style={{ ...inputStyle, width: '100%' }} value={cuentaUsernameEdit}
                    onChange={(e) => setCuentaUsernameEdit(e.target.value.toLowerCase().replace(/[^a-z0-9._]/g, ''))}
                    placeholder="ej. edgar.farias" />
                </Field>
                <div style={{ fontSize: 11, color: C.muted, marginTop: 4 }}>Solo minúsculas, números, puntos y guion bajo. Déjalo vacío si prefieres seguir usando tu correo.</div>
              </div>

              <div style={{ fontSize: 11, fontWeight: 700, color: C.muted, textTransform: 'uppercase', marginTop: 16 }}>Cambiar clave</div>
              <div className="flex gap-3 mt-2">
                <Field label="Nueva clave"><input type="password" style={inputStyle} value={cuentaClaveNueva} onChange={(e) => setCuentaClaveNueva(e.target.value)} /></Field>
                <Field label="Confirmar clave"><input type="password" style={inputStyle} value={cuentaClaveConfirm} onChange={(e) => setCuentaClaveConfirm(e.target.value)} /></Field>
              </div>

              {cuentaMsg && <div style={{ fontSize: 12, color: cuentaMsg.startsWith('✓') ? C.cyanDark : C.coral, marginTop: 10, fontWeight: 600 }}>{cuentaMsg}</div>}

              <button onClick={async () => {
                if (!cuentaNombreEdit.trim()) { setCuentaMsg('El nombre no puede quedar vacío.'); return; }
                if (cuentaClaveNueva && cuentaClaveNueva !== cuentaClaveConfirm) { setCuentaMsg('Las claves no coinciden.'); return; }
                const { error: errPerfil } = await supabase.from('perfiles').update({ nombre: cuentaNombreEdit.trim(), username: cuentaUsernameEdit.trim() || null }).eq('id', sesion.id);
                if (errPerfil) { setCuentaMsg(errPerfil.message.includes('duplicate') ? 'Ese usuario ya lo tiene alguien más — elige otro.' : 'No se pudo guardar: ' + errPerfil.message); return; }
                if (cuentaClaveNueva) {
                  const { error: errClave } = await supabase.auth.updateUser({ password: cuentaClaveNueva });
                  if (errClave) { setCuentaMsg('El nombre se guardó, pero la clave no: ' + errClave.message); return; }
                }
                setSesion((prev) => ({ ...prev, nombre: cuentaNombreEdit.trim(), username: cuentaUsernameEdit.trim() || null }));
                setCuentaClaveNueva(''); setCuentaClaveConfirm('');
                setCuentaMsg('✓ Cambios guardados.');
              }} className="mt-4 px-4 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.cyanDark }}>Guardar cambios</button>

              <div style={{ marginTop: 20, paddingTop: 16, borderTop: `1px solid ${C.border}` }}>
                <div style={{ fontSize: 11, fontWeight: 700, color: C.muted, textTransform: 'uppercase' }}>Acceso</div>
                <div style={{ fontSize: 12.5, color: C.slate, marginTop: 6 }}>Último acceso: {sesion.ultimoAcceso || 'este es tu primer ingreso'}</div>
                <div style={{ fontSize: 12.5, color: C.slate, marginTop: 3 }}>Sesiones registradas: {sesion.accesos || 1}</div>
                <div style={{ fontSize: 12.5, color: C.slate, marginTop: 3 }}>Empresas con acceso: {sesion.empresas === 'todas' ? 'Todas' : empresasDeSesion().map((e) => e.nombre).join(', ') || '—'}</div>
                {sesion.empleadoId && <div style={{ fontSize: 12.5, color: C.slate, marginTop: 3 }}>Vinculado al técnico: {empleadoName(sesion.empleadoId)}</div>}
              </div>
            </div>
          </div>
        )}

        {/* PREFERENCIAS (personal, cualquier usuario) */}
        {tab === 'preferencias' && (
          <div>
            <h1 style={{ fontFamily: 'Space Grotesk, sans-serif', fontSize: 22, fontWeight: 700, color: C.slate }}>{t('Preferencias')}</h1>
            <p style={{ color: C.muted, fontSize: 13, marginTop: 2 }}>Ajustes de cómo ves el sistema — son personales, no afectan a los demás usuarios.</p>

            <div className="rounded-xl p-5 mt-5" style={{ backgroundColor: C.card, border: `1px solid ${C.border}`, maxWidth: 480 }}>
              <Field label="Pantalla de inicio al entrar">
                <select style={{ ...inputStyle, width: '100%' }} value={sesion.preferencias?.pantallaInicio || 'dashboard'}
                  onChange={(e) => {
                    const val = e.target.value;
                    supabase.from('perfiles').update({ preferencias: { ...sesion.preferencias, pantallaInicio: val } }).eq('id', sesion.id).then(() => {});
                    setSesion((prev) => ({ ...prev, preferencias: { ...prev.preferencias, pantallaInicio: val } }));
                  }}>
                  <option value="dashboard">Panel general</option>
                  <option value="ordenes">Órdenes</option>
                  <option value="agenda">Agenda</option>
                  <option value="ventas">Ventas</option>
                  <option value="contab_informes">Informes</option>
                </select>
              </Field>

              <div style={{ marginTop: 16 }}>
                <Field label="Símbolo de moneda para mostrar montos">
                  <select style={{ ...inputStyle, width: '100%' }} value={sesion.preferencias?.moneda || '$'}
                    onChange={(e) => {
                      const val = e.target.value;
                      supabase.from('perfiles').update({ preferencias: { ...sesion.preferencias, moneda: val } }).eq('id', sesion.id).then(() => {});
                      setSesion((prev) => ({ ...prev, preferencias: { ...prev.preferencias, moneda: val } }));
                    }}>
                    <option value="$">$ (peso / dólar)</option>
                    <option value="€">€ (euro)</option>
                    <option value="£">£ (libra)</option>
                  </select>
                </Field>
                <div style={{ fontSize: 11, color: C.muted, marginTop: 6 }}>Solo cambia el símbolo mostrado — la contabilidad sigue en una sola moneda base.</div>
              </div>

              <div style={{ marginTop: 16 }}>
                <Field label="Idioma del sistema">
                  <select style={{ ...inputStyle, width: '100%' }} value={sesion.preferencias?.idioma || 'es'}
                    onChange={(e) => {
                      const val = e.target.value;
                      supabase.from('perfiles').update({ preferencias: { ...sesion.preferencias, idioma: val } }).eq('id', sesion.id).then(() => {});
                      setSesion((prev) => ({ ...prev, preferencias: { ...prev.preferencias, idioma: val } }));
                    }}>
                    <option value="es">Español</option>
                    <option value="it">Italiano</option>
                  </select>
                </Field>
                <div style={{ fontSize: 11, color: C.muted, marginTop: 6 }}>
                  Traduce el menú, los títulos de cada módulo y las acciones más comunes. El contenido dentro de cada pantalla (formularios, tablas) se sigue traduciendo por partes.
                </div>
              </div>

              <div style={{ marginTop: 20, paddingTop: 16, borderTop: `1px solid ${C.border}` }}>
                <div style={{ fontSize: 11, fontWeight: 700, color: C.muted, textTransform: 'uppercase' }}>Tema</div>
                <div style={{ fontSize: 12.5, color: C.muted, marginTop: 6 }}>El tema oscuro está en nuestra lista de mejoras — próximamente.</div>
              </div>
            </div>
          </div>
        )}

        {/* FORMATO DE CORREOS (plantillas — el envío real se activa al migrar a Supabase) */}
        {tab === 'formato_correos' && (
          <div>
            <h1 style={{ fontFamily: 'Space Grotesk, sans-serif', fontSize: 22, fontWeight: 700, color: C.slate }}>{t('Formato de correos')}</h1>
            <p style={{ color: C.muted, fontSize: 13, marginTop: 2 }}>Plantillas para los correos automáticos del sistema. Hoy se configuran y previsualizan aquí; el envío real se activa cuando migremos a Supabase.</p>

            <div className="flex gap-2 mt-4">
              {Object.keys(emailTemplates).map((k) => (
                <button key={k} onClick={() => setPlantillaSel(k)} className="px-3 py-1.5 rounded-full text-xs font-semibold" style={{ backgroundColor: plantillaSel === k ? C.navy : '#fff', color: plantillaSel === k ? '#fff' : C.muted, border: `1px solid ${C.border}` }}>
                  {emailTemplates[k].nombre}
                </button>
              ))}
            </div>

            {(() => {
              const pl = emailTemplates[plantillaSel];
              const vars = EMAIL_VARIABLES[plantillaSel] || [];
              const datosEjemplo = {
                '{squadra}': 'Squadra Norte', '{fecha}': new Date().toISOString().slice(0, 10), '{total_ordenes}': '4',
                '{lista_ordenes}': '• OT-0501 · 09:00 · Vía Roma 12\n• OT-0502 · 11:30 · Vía Milano 8',
                '{empresa}': empresaActual?.nombre || 'Tu empresa', '{cliente}': 'Marco Rossi', '{folio}': 'OT-0501',
                '{tecnico}': 'Carlos Medina', '{tipo_servicio}': 'Instalación residencial',
                '{vehiculo}': 'Camioneta 1', '{placa}': 'ABC-123', '{tipo_vencimiento}': 'Seguro', '{fecha_vencimiento}': '2026-08-15', '{dias_restantes}': '18',
              };
              const previsualizar = (texto) => vars.reduce((t, v) => t.split(v).join(datosEjemplo[v] || v), texto);
              return (
                <div className="grid grid-cols-2 gap-5 mt-4">
                  <div className="rounded-xl p-5" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
                    <div style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 600, fontSize: 14, color: C.slate, marginBottom: 10 }}>Editar plantilla</div>
                    <Field label="Asunto"><input style={{ ...inputStyle, width: '100%' }} value={pl.asunto} onChange={(e) => setEmailTemplates((prev) => ({ ...prev, [plantillaSel]: { ...prev[plantillaSel], asunto: e.target.value } }))} /></Field>
                    <div className="mt-3">
                      <div style={{ fontSize: 12, color: C.muted, marginBottom: 4 }}>Cuerpo del mensaje</div>
                      <textarea value={pl.cuerpo} onChange={(e) => setEmailTemplates((prev) => ({ ...prev, [plantillaSel]: { ...prev[plantillaSel], cuerpo: e.target.value } }))}
                        style={{ ...inputStyle, width: '100%', minHeight: 160, fontFamily: 'Inter, sans-serif', resize: 'vertical' }} />
                    </div>
                    <div style={{ fontSize: 11, color: C.muted, marginTop: 10 }}>Variables disponibles: {vars.map((v) => <span key={v} style={{ fontFamily: 'JetBrains Mono, monospace', backgroundColor: '#EEF1F4', padding: '1px 5px', borderRadius: 4, marginRight: 4 }}>{v}</span>)}</div>
                  </div>
                  <div className="rounded-xl p-5" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
                    <div style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 600, fontSize: 14, color: C.slate, marginBottom: 10 }}>Vista previa (con datos de ejemplo)</div>
                    <div className="rounded-lg p-3" style={{ backgroundColor: '#F6F8FA', border: `1px solid ${C.border}` }}>
                      <div style={{ fontSize: 11, color: C.muted }}>Asunto</div>
                      <div style={{ fontSize: 13, fontWeight: 600, color: C.slate, marginBottom: 10 }}>{previsualizar(pl.asunto)}</div>
                      <div style={{ fontSize: 11, color: C.muted }}>Cuerpo</div>
                      <div style={{ fontSize: 12.5, color: C.slate, whiteSpace: 'pre-wrap', marginTop: 4 }}>{previsualizar(pl.cuerpo)}</div>
                    </div>
                  </div>
                </div>
              );
            })()}
          </div>
        )}

        {/* INGRESOS Y GASTOS GENERALES */}
        {tab === 'ingresos_gastos' && (
          <div>
            <div className="flex items-center justify-between">
              <div>
                <h1 style={{ fontFamily: 'Space Grotesk, sans-serif', fontSize: 22, fontWeight: 700, color: C.slate }}>{t('Ingresos y gastos generales')}</h1>
                <p style={{ color: C.muted, fontSize: 13, marginTop: 2 }}>Movimientos de la empresa que no vienen de una orden: renta, nómina, combustible, ventas de activos, etc.</p>
              </div>
              <button onClick={() => setShowMovForm(!showMovForm)} className="flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.navy }}>
                {showMovForm ? <X size={15} /> : <Plus size={15} />} {showMovForm ? 'Cerrar' : 'Nuevo movimiento'}
              </button>
            </div>

            <div className="grid grid-cols-3 gap-4 mt-5">
              <KPICard label="Ingresos adicionales" value={fmt(ingresosAdicionales)} accent={C.cyanDark} />
              <KPICard label="Gastos generales" value={fmt(gastosAdicionales)} accent={C.coral} />
              <KPICard label="Balance adicional" value={fmt(ingresosAdicionales - gastosAdicionales)} accent={ingresosAdicionales - gastosAdicionales >= 0 ? C.cyanDark : C.coral} />
            </div>

            {showMovForm && <MovimientoForm gastoCategorias={gastoCategorias} ingresoCategorias={ingresoCategorias} onSubmit={(data) => { addMovimiento(data); setShowMovForm(false); }} />}

            <div className="flex gap-1.5 mt-5 mb-3">
              {['todos', 'ingreso', 'gasto'].map((f) => (
                <button key={f} onClick={() => setMovFilter(f)} className="px-3 py-1.5 rounded-full text-xs font-semibold"
                  style={{ backgroundColor: movFilter === f ? C.navy : '#fff', color: movFilter === f ? '#fff' : C.muted, border: `1px solid ${C.border}` }}>
                  {f === 'todos' ? 'Todos' : f === 'ingreso' ? 'Ingresos' : 'Gastos'}
                </button>
              ))}
            </div>

            <div className="rounded-xl overflow-hidden" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
              <table>
                <thead><tr><th></th><th>Fecha</th><th>Categoría</th><th>Descripción</th><th>Monto</th><th></th></tr></thead>
                <tbody>
                  {movimientos.filter((m) => movFilter === 'todos' || m.tipo === movFilter).sort((a, b) => b.fecha.localeCompare(a.fecha)).map((m) => (
                    <tr key={m.id}>
                      <td>{m.tipo === 'ingreso' ? <ArrowUpCircle size={16} color={C.cyanDark} /> : <ArrowDownCircle size={16} color={C.coral} />}</td>
                      <td>{m.fecha}</td>
                      <td style={{ fontSize: 12.5 }}>{m.categoria}</td>
                      <td>{m.descripcion}</td>
                      <td style={{ fontFamily: 'JetBrains Mono, monospace', color: m.tipo === 'ingreso' ? C.cyanDark : C.coral }}>{m.tipo === 'ingreso' ? '+' : '−'}{fmt(m.monto)}</td>
                      <td><button onClick={() => deleteMovimiento(m.id)}><Trash2 size={14} color={C.coral} /></button></td>
                    </tr>
                  ))}
                  {movimientos.length === 0 && <tr><td colSpan={6} style={{ color: C.muted, textAlign: 'center', padding: 24 }}>Sin movimientos registrados.</td></tr>}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* REPORTES */}
        {tab === 'reportes' && (
          <div>
            <div className="flex items-center justify-between">
              <div>
                <h1 style={{ fontFamily: 'Space Grotesk, sans-serif', fontSize: 22, fontWeight: 700, color: C.slate }}>{t('Reportes')}</h1>
                <p style={{ color: C.muted, fontSize: 13, marginTop: 2 }}>Elige el tipo de reporte que necesitas.</p>
              </div>
              <button onClick={exportReporte} className="flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-sm font-semibold" style={{ backgroundColor: '#fff', color: C.slate, border: `1px solid ${C.border}` }}>
                <Download size={15} /> Exportar CSV
              </button>
            </div>

            <div className="flex gap-3 items-end mt-5">
              <Field label="Tipo de reporte">
                <select style={{ ...inputStyle, width: 240 }} value={reportType} onChange={(e) => setReportType(e.target.value)}>
                  {REPORT_TYPES.map((r) => <option key={r.key} value={r.key}>{r.label}</option>)}
                </select>
              </Field>
              <Field label="Del"><input type="date" style={inputStyle} value={reportFrom} onChange={(e) => setReportFrom(e.target.value)} /></Field>
              <Field label="Al"><input type="date" style={inputStyle} value={reportTo} onChange={(e) => setReportTo(e.target.value)} /></Field>
              {['cxc', 'cxp', 'inventario', 'activos'].includes(reportType) && (
                <span style={{ fontSize: 11.5, color: C.muted, marginBottom: 8 }}>Este reporte muestra el corte actual; el rango de fechas no aplica.</span>
              )}
            </div>

            {/* ESTADO DE RESULTADOS */}
            {reportType === 'estado_resultados' && (
              <div className="rounded-xl p-5 mt-6" style={{ backgroundColor: C.card, border: `1px solid ${C.border}`, maxWidth: 560 }}>
                <div style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 600, fontSize: 14, color: C.slate, marginBottom: 10 }}>Estado de resultados</div>
                {[
                  ['Ingresos por instalaciones', repIngresosOrdenes, false],
                  ['(-) Costo de materiales', -repCostoMateriales, false],
                  ['(-) Costo de mano de obra', -repCostoManoObra, false],
                  ['= Utilidad operativa', repUtilidadOperativa, true],
                  ['(+) Otros ingresos', repIngresosAd, false],
                  ['(-) Gastos generales', -repGastosAd, false],
                  ['= Utilidad neta del periodo', repUtilidadNeta, true],
                ].map(([label, val, strong], i) => (
                  <div key={i} className="flex justify-between py-1.5" style={{ borderTop: strong ? `1px solid ${C.border}` : 'none', marginTop: strong ? 6 : 0 }}>
                    <span style={{ fontSize: 13, fontWeight: strong ? 700 : 400, color: C.slate }}>{label}</span>
                    <span style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 13, fontWeight: strong ? 700 : 500, color: val >= 0 ? (strong ? C.cyanDark : C.slate) : C.coral }}>{fmt(val)}</span>
                  </div>
                ))}
              </div>
            )}

            {/* VENTAS */}
            {reportType === 'ventas' && (
              <div className="rounded-xl overflow-hidden mt-6" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
                <table>
                  <thead><tr><th>Folio</th><th>Cliente</th><th>Proyecto</th><th>Fecha</th><th>Precio</th><th>Estado de pago</th></tr></thead>
                  <tbody>
                    {repOrdenes.map((o) => (
                      <tr key={o.id}>
                        <td style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 12 }}>{o.id}</td>
                        <td>{clientName(o.clienteId)}</td>
                        <td style={{ fontSize: 12 }}>{o.proyectoId ? projectName(o.proyectoId) : '—'}</td>
                        <td>{o.fecha}</td>
                        <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(o.precioCliente)}</td>
                        <td><SignalBadge estado={o.estadoPago === 'pagada' ? 'completada' : 'pendiente'} /></td>
                      </tr>
                    ))}
                    {repOrdenes.length === 0 && <tr><td colSpan={6} style={{ color: C.muted, textAlign: 'center', padding: 24 }}>Sin ventas en el periodo.</td></tr>}
                  </tbody>
                  <tfoot><tr><td colSpan={4} style={{ fontWeight: 700 }}>Total</td><td style={{ fontFamily: 'JetBrains Mono, monospace', fontWeight: 700 }}>{fmt(repIngresosOrdenes)}</td><td></td></tr></tfoot>
                </table>
              </div>
            )}

            {/* COSTOS POR ORDEN */}
            {reportType === 'costos' && (
              <div className="rounded-xl overflow-hidden mt-6" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
                <table>
                  <thead><tr><th>Folio</th><th>Costo materiales</th><th>Costo mano de obra</th><th>Costo total</th><th>Utilidad</th></tr></thead>
                  <tbody>
                    {repOrdenes.map((o) => (
                      <tr key={o.id}>
                        <td style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 12 }}>{o.id}</td>
                        <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(materialCost(o))}</td>
                        <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(o.costoManoObra)}</td>
                        <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(orderCost(o))}</td>
                        <td style={{ fontFamily: 'JetBrains Mono, monospace', color: orderProfit(o) >= 0 ? C.cyanDark : C.coral }}>{fmt(orderProfit(o))}</td>
                      </tr>
                    ))}
                    {repOrdenes.length === 0 && <tr><td colSpan={5} style={{ color: C.muted, textAlign: 'center', padding: 24 }}>Sin órdenes completadas en el periodo.</td></tr>}
                  </tbody>
                  <tfoot><tr><td style={{ fontWeight: 700 }}>Total</td><td style={{ fontFamily: 'JetBrains Mono, monospace', fontWeight: 700 }}>{fmt(repCostoMateriales)}</td><td style={{ fontFamily: 'JetBrains Mono, monospace', fontWeight: 700 }}>{fmt(repCostoManoObra)}</td><td style={{ fontFamily: 'JetBrains Mono, monospace', fontWeight: 700 }}>{fmt(repCostoMateriales + repCostoManoObra)}</td><td></td></tr></tfoot>
                </table>
              </div>
            )}

            {/* GASTOS POR CATEGORIA */}
            {reportType === 'gastos_categoria' && (
              <div className="rounded-xl p-5 mt-6" style={{ backgroundColor: C.card, border: `1px solid ${C.border}`, maxWidth: 560 }}>
                <div style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 600, fontSize: 14, color: C.slate, marginBottom: 10 }}>Gastos generales por categoría</div>
                {gastosPorCategoria.length === 0 && <div style={{ color: C.muted, fontSize: 13 }}>Sin gastos en el periodo seleccionado.</div>}
                {gastosPorCategoria.map((c) => (
                  <div key={c.categoria} className="mb-2.5">
                    <div className="flex justify-between text-xs mb-1"><span style={{ color: C.slate }}>{c.categoria}</span><span style={{ fontFamily: 'JetBrains Mono, monospace', color: C.muted }}>{fmt(c.monto)}</span></div>
                    <div style={{ height: 6, borderRadius: 3, backgroundColor: C.bg, overflow: 'hidden' }}>
                      <div style={{ height: '100%', width: `${(c.monto / maxGastoCat) * 100}%`, backgroundColor: C.coral, borderRadius: 3 }} />
                    </div>
                  </div>
                ))}
              </div>
            )}

            {/* CUENTAS POR COBRAR */}
            {reportType === 'cxc' && (
              <div className="rounded-xl overflow-hidden mt-6" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
                <table>
                  <thead><tr><th>Folio</th><th>Cliente</th><th>Fecha</th><th>Monto</th></tr></thead>
                  <tbody>
                    {completedOrders.filter((o) => o.estadoPago === 'pendiente').map((o) => (
                      <tr key={o.id}>
                        <td style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 12 }}>{o.id}</td>
                        <td>{clientName(o.clienteId)}</td>
                        <td>{o.fecha}</td>
                        <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(o.precioCliente)}</td>
                      </tr>
                    ))}
                    {completedOrders.filter((o) => o.estadoPago === 'pendiente').length === 0 && <tr><td colSpan={4} style={{ color: C.muted, textAlign: 'center', padding: 24 }}>No hay cuentas por cobrar pendientes.</td></tr>}
                  </tbody>
                  <tfoot><tr><td colSpan={3} style={{ fontWeight: 700 }}>Total por cobrar</td><td style={{ fontFamily: 'JetBrains Mono, monospace', fontWeight: 700 }}>{fmt(receivable)}</td></tr></tfoot>
                </table>
              </div>
            )}

            {/* CUENTAS POR PAGAR */}
            {reportType === 'cxp' && (
              <div className="rounded-xl overflow-hidden mt-6" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
                <table>
                  <thead><tr><th>Proveedor</th><th>Concepto</th><th>Vencimiento</th><th>Monto</th><th>Estado</th></tr></thead>
                  <tbody>
                    {[...cxp].sort((a, b) => a.vencimiento.localeCompare(b.vencimiento)).map((c) => (
                      <tr key={c.id}>
                        <td>{proveedorName(c.proveedorId)}</td>
                        <td>{c.concepto}</td>
                        <td>{c.vencimiento}</td>
                        <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(c.monto)}</td>
                        <td><SignalBadge estado={c.estado === 'pagada' ? 'completada' : 'pendiente'} /></td>
                      </tr>
                    ))}
                  </tbody>
                  <tfoot><tr><td colSpan={3} style={{ fontWeight: 700 }}>Total pendiente</td><td style={{ fontFamily: 'JetBrains Mono, monospace', fontWeight: 700 }}>{fmt(cxpPendiente)}</td><td></td></tr></tfoot>
                </table>
              </div>
            )}

            {/* INVENTARIO VALORIZADO */}
            {reportType === 'inventario' && (
              <div className="rounded-xl overflow-hidden mt-6" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
                <table>
                  <thead><tr><th>Material</th><th>Stock</th><th>Costo promedio</th><th>Valor</th></tr></thead>
                  <tbody>
                    {materials.map((m) => (
                      <tr key={m.id}>
                        <td>{m.nombre}</td>
                        <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{m.stock}</td>
                        <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(m.costoUnitario)}</td>
                        <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(m.stock * m.costoUnitario)}</td>
                      </tr>
                    ))}
                  </tbody>
                  <tfoot><tr><td colSpan={3} style={{ fontWeight: 700 }}>Valor total de inventario</td><td style={{ fontFamily: 'JetBrains Mono, monospace', fontWeight: 700 }}>{fmt(inventoryValue)}</td></tr></tfoot>
                </table>
              </div>
            )}

            {/* NOMINA */}
            {reportType === 'nomina' && (
              <div className="rounded-xl overflow-hidden mt-6" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
                <table>
                  <thead><tr><th>Empleado</th><th>Periodo</th><th>Fecha de pago</th><th>Monto</th><th>Estado</th></tr></thead>
                  <tbody>
                    {nomina.filter((n) => inRange(n.fechaPago)).map((n) => (
                      <tr key={n.id}>
                        <td>{empleadoName(n.empleadoId)}</td>
                        <td>{n.periodo}</td>
                        <td>{n.fechaPago}</td>
                        <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(n.monto)}</td>
                        <td><SignalBadge estado={n.estado === 'pagado' ? 'completada' : 'pendiente'} /></td>
                      </tr>
                    ))}
                    {nomina.filter((n) => inRange(n.fechaPago)).length === 0 && <tr><td colSpan={5} style={{ color: C.muted, textAlign: 'center', padding: 24 }}>Sin nómina en el periodo.</td></tr>}
                  </tbody>
                </table>
              </div>
            )}

            {/* ACTIVOS FIJOS */}
            {reportType === 'activos' && (
              <div className="rounded-xl overflow-hidden mt-6" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
                <table>
                  <thead><tr><th>Activo</th><th>Categoría</th><th>Centro de costo</th><th>Costo</th><th>Dep. mensual</th><th>Dep. contabilizada</th><th>Valor en libros contable</th></tr></thead>
                  <tbody>
                    {activos.map((a) => {
                      const d = depreciacionInfo(a);
                      const contab = depContabilizada(a.id);
                      return (
                        <tr key={a.id}>
                          <td>{a.nombre}</td>
                          <td>{a.categoria}</td>
                          <td style={{ fontSize: 12 }}>{centroCostoName(a.centroCostoId)}</td>
                          <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(a.costo)}</td>
                          <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(d.depAnual / 12)}</td>
                          <td style={{ fontFamily: 'JetBrains Mono, monospace', color: C.amber }}>{fmt(contab)}</td>
                          <td style={{ fontFamily: 'JetBrains Mono, monospace', color: C.cyanDark }}>{fmt(a.costo - contab)}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                  <tfoot><tr><td colSpan={6} style={{ fontWeight: 700 }}>Valor en libros estimado total</td><td style={{ fontFamily: 'JetBrains Mono, monospace', fontWeight: 700 }}>{fmt(valorActivosLibros)}</td></tr></tfoot>
                </table>
              </div>
            )}

            {/* UTILIDAD POR PROYECTO */}
            {reportType === 'proyectos' && (
              <div className="rounded-xl overflow-hidden mt-6" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
                <table>
                  <thead><tr><th>Proyecto</th><th>Cliente</th><th>Presupuesto</th><th>Gastado</th><th>Facturado</th><th>Utilidad</th></tr></thead>
                  <tbody>
                    {projects.map((p) => {
                      const m = projectMetrics(p);
                      return (
                        <tr key={p.id}>
                          <td>{p.nombre}</td>
                          <td>{clientName(p.clienteId)}</td>
                          <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(p.presupuesto)}</td>
                          <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(m.gastado)}</td>
                          <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(m.ingresos)}</td>
                          <td style={{ fontFamily: 'JetBrains Mono, monospace', color: m.utilidad >= 0 ? C.cyanDark : C.coral }}>{fmt(m.utilidad)}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}

            {/* RENTABILIDAD POR TECNICO */}
            {reportType === 'tecnicos' && (
              <div className="rounded-xl overflow-hidden mt-6" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
                <table>
                  <thead><tr><th>Técnico</th><th>Órdenes completadas</th><th>Mano de obra generada</th><th>Pendiente por pagar</th></tr></thead>
                  <tbody>
                    {techs.map((t) => {
                      const done = orders.filter((o) => o.tecnicoId === t.id && o.estado === 'completada');
                      const ganado = done.reduce((s, o) => s + Number(o.costoManoObra), 0);
                      const pendiente = done.filter((o) => o.estadoPagoTecnico === 'pendiente').reduce((s, o) => s + Number(o.costoManoObra), 0);
                      return (
                        <tr key={t.id}>
                          <td>{t.nombre}</td>
                          <td>{done.length}</td>
                          <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(ganado)}</td>
                          <td style={{ fontFamily: 'JetBrains Mono, monospace', color: pendiente > 0 ? C.amber : C.muted }}>{fmt(pendiente)}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}

            {/* CENTROS DE COSTO */}
            {reportType === 'centros_costo' && (
              <div className="mt-6">
                <div className="rounded-xl overflow-hidden" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
                  <div className="px-4 pt-4 pb-1" style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 600, fontSize: 14, color: C.slate }}>Consolidado global</div>
                  <table>
                    <thead><tr><th>Centro de costo</th><th>Órdenes</th><th>Ingresos</th><th>Costo materiales</th><th>Costo mano de obra</th><th>Utilidad</th></tr></thead>
                    <tbody>
                      {centros.map((c) => {
                        const m = centroCostoMetrics(c.id, repOrdenes);
                        return (
                          <tr key={c.id}>
                            <td>{c.nombre}</td>
                            <td>{m.ordenes}</td>
                            <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(m.ingresos)}</td>
                            <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(m.cMateriales)}</td>
                            <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(m.cManoObra)}</td>
                            <td style={{ fontFamily: 'JetBrains Mono, monospace', color: m.utilidad >= 0 ? C.cyanDark : C.coral }}>{fmt(m.utilidad)}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                    <tfoot>
                      <tr>
                        <td style={{ fontWeight: 700 }}>Global</td>
                        <td style={{ fontWeight: 700 }}>{repOrdenes.length}</td>
                        <td style={{ fontFamily: 'JetBrains Mono, monospace', fontWeight: 700 }}>{fmt(repIngresosOrdenes)}</td>
                        <td style={{ fontFamily: 'JetBrains Mono, monospace', fontWeight: 700 }}>{fmt(repCostoMateriales)}</td>
                        <td style={{ fontFamily: 'JetBrains Mono, monospace', fontWeight: 700 }}>{fmt(repCostoManoObra)}</td>
                        <td style={{ fontFamily: 'JetBrains Mono, monospace', fontWeight: 700, color: repUtilidadOperativa >= 0 ? C.cyanDark : C.coral }}>{fmt(repUtilidadOperativa)}</td>
                      </tr>
                    </tfoot>
                  </table>
                </div>

                <div className="grid grid-cols-3 gap-4 mt-5">
                  {centros.map((c) => {
                    const m = centroCostoMetrics(c.id, repOrdenes);
                    const relacionados = repOrdenes.filter((o) => o.centroCostoId === c.id);
                    return (
                      <div key={c.id} className="rounded-xl overflow-hidden" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
                        <div className="px-4 pt-4 pb-2">
                          <div style={{ fontFamily: 'Space Grotesk, sans-serif', fontWeight: 600, fontSize: 14, color: C.slate }}>{c.nombre}</div>
                          <div style={{ fontSize: 11.5, color: C.muted, marginTop: 2 }}>Utilidad: <span style={{ color: m.utilidad >= 0 ? C.cyanDark : C.coral, fontWeight: 600 }}>{fmt(m.utilidad)}</span></div>
                        </div>
                        <table>
                          <thead><tr><th>Folio</th><th>Precio</th></tr></thead>
                          <tbody>
                            {relacionados.map((o) => (
                              <tr key={o.id}><td style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 11.5 }}>{o.id}</td><td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{fmt(o.precioCliente)}</td></tr>
                            ))}
                            {relacionados.length === 0 && <tr><td colSpan={2} style={{ color: C.muted, textAlign: 'center', padding: 14 }}>Sin órdenes en el periodo.</td></tr>}
                          </tbody>
                        </table>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}
          </div>
        )}
      </div>
      </div>
    </div>
  );
}

function NominaForm({ empleados, onSubmit }) {
  const [empleadoId, setEmpleadoId] = useState(empleados[0]?.id || '');
  const [periodo, setPeriodo] = useState('');
  const [fechaPago, setFechaPago] = useState(new Date().toISOString().slice(0, 10));
  const [monto, setMonto] = useState(empleados[0]?.salarioMensual || '');
  return (
    <div className="rounded-xl p-4 mt-4 flex flex-wrap gap-3 items-end" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
      <Field label="Empleado">
        <select style={{ ...inputStyle, width: 200 }} value={empleadoId} onChange={(e) => { setEmpleadoId(e.target.value); const emp = empleados.find((x) => x.id === e.target.value); if (emp) setMonto(emp.salarioMensual); }}>
          {empleados.map((e) => <option key={e.id} value={e.id}>{e.nombre}</option>)}
        </select>
      </Field>
      <Field label="Periodo (ej. Julio 2026)"><input style={{ ...inputStyle, width: 160 }} value={periodo} onChange={(e) => setPeriodo(e.target.value)} /></Field>
      <Field label="Fecha de pago"><input type="date" style={inputStyle} value={fechaPago} onChange={(e) => setFechaPago(e.target.value)} /></Field>
      <Field label="Monto ($)"><input type="number" style={{ ...inputStyle, width: 120 }} value={monto} onChange={(e) => setMonto(e.target.value)} /></Field>
      <button onClick={() => { if (empleadoId && periodo && monto !== '') onSubmit({ empleadoId, periodo, fechaPago, monto }); }} className="px-4 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.cyanDark }}>Registrar</button>
      <div style={{ fontSize: 11, color: C.muted, width: '100%' }}>Esto genera automáticamente el asiento contable del gasto de nómina. Al marcarla "Pagado" se registra la salida de caja.</div>
    </div>
  );
}

function CxpForm({ proveedores, gastoCategorias, materials, onSubmit }) {
  const [proveedorId, setProveedorId] = useState(proveedores[0]?.id || '');
  const [concepto, setConcepto] = useState('');
  const [tipo, setTipo] = useState('gasto');
  const [categoria, setCategoria] = useState(gastoCategorias[0]?.nombre || '');
  const [materialId, setMaterialId] = useState(materials[0]?.id || '');
  const [cantidadLinea, setCantidadLinea] = useState('');
  const [montoLinea, setMontoLinea] = useState('');
  const [lineas, setLineas] = useState([]); // [{materialId, nombre, cantidad, monto}]
  const [fecha, setFecha] = useState(new Date().toISOString().slice(0, 10));
  const [vencimiento, setVencimiento] = useState(new Date().toISOString().slice(0, 10));
  const [monto, setMonto] = useState(''); // solo se usa cuando tipo === 'gasto'
  const totalLineas = lineas.reduce((s, l) => s + Number(l.monto), 0);

  const agregarLinea = () => {
    if (!materialId || cantidadLinea === '' || montoLinea === '') return;
    const m = materials.find((x) => x.id === materialId);
    setLineas((prev) => [...prev, { materialId, nombre: m?.nombre || '—', cantidad: Number(cantidadLinea), monto: Number(montoLinea) }]);
    setCantidadLinea(''); setMontoLinea('');
  };

  return (
    <div className="rounded-xl p-4 mt-4" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
      <div className="flex gap-2 mb-3">
        {[['gasto', 'Es un gasto'], ['inventario', 'Es compra de inventario']].map(([k, l]) => (
          <button key={k} onClick={() => setTipo(k)} className="px-3 py-1.5 rounded-full text-xs font-semibold" style={{ backgroundColor: tipo === k ? C.navy : '#fff', color: tipo === k ? '#fff' : C.muted, border: `1px solid ${C.border}` }}>{l}</button>
        ))}
      </div>
      <div className="flex flex-wrap gap-3 items-end">
        <Field label="Proveedor"><select style={{ ...inputStyle, width: 200 }} value={proveedorId} onChange={(e) => setProveedorId(e.target.value)}>{proveedores.map((p) => <option key={p.id} value={p.id}>{p.nombre}</option>)}</select></Field>
        <Field label="Concepto (factura / folio)"><input style={{ ...inputStyle, width: 200 }} value={concepto} onChange={(e) => setConcepto(e.target.value)} /></Field>
        {tipo === 'gasto' && (
          <>
            <Field label="Categoría de gasto">
              <select style={{ ...inputStyle, width: 200 }} value={categoria} onChange={(e) => setCategoria(e.target.value)}>
                {gastoCategorias.map((c) => <option key={c.nombre} value={c.nombre}>{c.nombre}</option>)}
              </select>
            </Field>
            <Field label="Monto ($)"><input type="number" style={{ ...inputStyle, width: 120 }} value={monto} onChange={(e) => setMonto(e.target.value)} /></Field>
          </>
        )}
        <Field label="Fecha"><input type="date" style={inputStyle} value={fecha} onChange={(e) => setFecha(e.target.value)} /></Field>
        <Field label="Vencimiento"><input type="date" style={inputStyle} value={vencimiento} onChange={(e) => setVencimiento(e.target.value)} /></Field>
      </div>

      {tipo === 'inventario' && (
        <div className="mt-4">
          <div style={{ fontSize: 11, color: C.muted, marginBottom: 8 }}>Agrega tantos materiales como traiga esta factura — cada uno con su cantidad y su costo total en esta compra.</div>
          <div className="flex flex-wrap gap-2 items-end">
            <Field label="Material">
              <select style={{ ...inputStyle, width: 200 }} value={materialId} onChange={(e) => setMaterialId(e.target.value)}>
                {materials.map((m) => <option key={m.id} value={m.id}>{m.nombre} (stock: {m.stock})</option>)}
              </select>
            </Field>
            <Field label="Cantidad"><input type="number" style={{ ...inputStyle, width: 90 }} value={cantidadLinea} onChange={(e) => setCantidadLinea(e.target.value)} /></Field>
            <Field label="Costo total ($)"><input type="number" style={{ ...inputStyle, width: 110 }} value={montoLinea} onChange={(e) => setMontoLinea(e.target.value)} /></Field>
            <button onClick={agregarLinea} className="px-3 py-2 rounded-lg text-sm font-semibold" style={{ backgroundColor: '#EEF1F4', color: C.slate }}>+ Añadir material</button>
          </div>

          {lineas.length > 0 && (
            <div className="rounded-lg overflow-hidden mt-3" style={{ border: `1px solid ${C.border}` }}>
              <table style={{ width: '100%' }}>
                <thead><tr style={{ backgroundColor: '#F7F8FA' }}><th style={{ fontSize: 10.5 }}>Material</th><th style={{ fontSize: 10.5, width: 80 }}>Cantidad</th><th style={{ fontSize: 10.5, width: 100 }}>Costo</th><th style={{ width: 30 }}></th></tr></thead>
                <tbody>
                  {lineas.map((l, i) => (
                    <tr key={i}>
                      <td style={{ padding: '5px 8px', fontSize: 12 }}>{l.nombre}</td>
                      <td style={{ padding: '5px 8px', fontSize: 12, textAlign: 'right' }}>{l.cantidad}</td>
                      <td style={{ padding: '5px 8px', fontSize: 12, textAlign: 'right', fontFamily: 'JetBrains Mono, monospace' }}>{fmt(l.monto)}</td>
                      <td style={{ padding: '5px 8px', textAlign: 'center' }}><button onClick={() => setLineas((prev) => prev.filter((_, idx) => idx !== i))}><X size={12} color={C.coral} /></button></td>
                    </tr>
                  ))}
                </tbody>
                <tfoot><tr style={{ borderTop: `1px solid ${C.border}` }}><td style={{ padding: '5px 8px', fontWeight: 700, fontSize: 12 }}>Total</td><td></td><td style={{ padding: '5px 8px', fontWeight: 700, fontSize: 12, textAlign: 'right', fontFamily: 'JetBrains Mono, monospace' }}>{fmt(totalLineas)}</td><td></td></tr></tfoot>
              </table>
            </div>
          )}
          <div style={{ fontSize: 11, color: C.muted, marginTop: 8 }}>Al guardar, cada material sube su stock y recalcula su costo promedio ponderado — se contabiliza a Inventario (1435), no como gasto.</div>
        </div>
      )}

      <button onClick={() => {
        if (!concepto) return;
        if (tipo === 'gasto' && monto !== '') onSubmit({ proveedorId, concepto, tipo, categoria, fecha, vencimiento, monto });
        else if (tipo === 'inventario' && lineas.length > 0) onSubmit({ proveedorId, concepto, tipo, fecha, vencimiento, monto: totalLineas, lineas });
      }} className="mt-4 px-4 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.cyanDark }}>Guardar</button>
    </div>
  );
}

function ActivoForm({ techs, centros, onSubmit }) {
  const [nombre, setNombre] = useState('');
  const [categoria, setCategoria] = useState('Vehículo');
  const [fechaCompra, setFechaCompra] = useState(new Date().toISOString().slice(0, 10));
  const [costo, setCosto] = useState('');
  const [valorResidual, setValorResidual] = useState('0');
  const [vidaUtilAnios, setVidaUtilAnios] = useState('5');
  const [responsable, setResponsable] = useState(techs[0]?.id || '');
  const [centroCostoId, setCentroCostoId] = useState(centros[0]?.id || '');
  const [placa, setPlaca] = useState('');
  const [km, setKm] = useState('');
  const [seguroVencimiento, setSeguroVencimiento] = useState('');
  const [revisionVencimiento, setRevisionVencimiento] = useState('');
  return (
    <div className="rounded-xl p-4 mt-4 flex flex-wrap gap-3 items-end" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
      <Field label="Nombre del activo"><input style={{ ...inputStyle, width: 220 }} value={nombre} onChange={(e) => setNombre(e.target.value)} /></Field>
      <Field label="Categoría">
        <select style={inputStyle} value={categoria} onChange={(e) => setCategoria(e.target.value)}>
          {['Vehículo', 'Equipo', 'Herramienta', 'Mobiliario', 'Otro'].map((c) => <option key={c}>{c}</option>)}
        </select>
      </Field>
      {categoria === 'Vehículo' && (
        <>
          <Field label="Placa"><input style={{ ...inputStyle, width: 110, fontFamily: 'JetBrains Mono, monospace' }} value={placa} onChange={(e) => setPlaca(e.target.value.toUpperCase())} /></Field>
          <Field label="Km actual"><input type="number" style={{ ...inputStyle, width: 100 }} value={km} onChange={(e) => setKm(e.target.value)} /></Field>
          <Field label="Vence seguro"><input type="date" style={inputStyle} value={seguroVencimiento} onChange={(e) => setSeguroVencimiento(e.target.value)} /></Field>
          <Field label="Vence revisión"><input type="date" style={inputStyle} value={revisionVencimiento} onChange={(e) => setRevisionVencimiento(e.target.value)} /></Field>
        </>
      )}
      <Field label="Responsable"><select style={inputStyle} value={responsable} onChange={(e) => setResponsable(e.target.value)}>{techs.map((t) => <option key={t.id} value={t.id}>{t.nombre}</option>)}</select></Field>
      <Field label="Centro de costo">
        <select style={inputStyle} value={centroCostoId} onChange={(e) => setCentroCostoId(e.target.value)}>
          {centros.map((c) => <option key={c.id} value={c.id}>{c.nombre}</option>)}
        </select>
      </Field>
      <Field label="Fecha de compra"><input type="date" style={inputStyle} value={fechaCompra} onChange={(e) => setFechaCompra(e.target.value)} /></Field>
      <Field label="Costo ($)"><input type="number" style={{ ...inputStyle, width: 110 }} value={costo} onChange={(e) => setCosto(e.target.value)} /></Field>
      <Field label="Valor residual ($)"><input type="number" style={{ ...inputStyle, width: 110 }} value={valorResidual} onChange={(e) => setValorResidual(e.target.value)} /></Field>
      <Field label="Vida útil (años)"><input type="number" style={{ ...inputStyle, width: 90 }} value={vidaUtilAnios} onChange={(e) => setVidaUtilAnios(e.target.value)} /></Field>
      <button onClick={() => { if (nombre && costo !== '') onSubmit({ nombre, categoria, fechaCompra, costo, valorResidual, vidaUtilAnios, responsable, centroCostoId, placa: placa || null, km: km ? Number(km) : null, seguroVencimiento: seguroVencimiento || null, revisionVencimiento: revisionVencimiento || null }); }} className="px-4 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.cyanDark }}>Guardar</button>
    </div>
  );
}

function FlotaForm({ activoNombre, onSubmit, onClose }) {
  const [tipo, setTipo] = useState('combustible');
  const [fecha, setFecha] = useState(new Date().toISOString().slice(0, 10));
  const [monto, setMonto] = useState('');
  const [kmActual, setKmActual] = useState('');
  const [subtipo, setSubtipo] = useState('ordinaria');
  const [detalle, setDetalle] = useState('');
  return (
    <div className="rounded-xl p-4 mt-4" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
      <div className="flex items-center justify-between">
        <div style={{ fontSize: 13, fontWeight: 600, color: C.slate }}>Nuevo registro — {activoNombre}</div>
        <button onClick={onClose}><X size={15} color={C.muted} /></button>
      </div>
      <div className="flex gap-2 mt-3">
        {[['combustible', 'Combustible'], ['mantenimiento', 'Mantenimiento'], ['multa', 'Multa']].map(([k, l]) => (
          <button key={k} onClick={() => setTipo(k)} className="px-3 py-1.5 rounded-full text-xs font-semibold" style={{ backgroundColor: tipo === k ? C.navy : '#fff', color: tipo === k ? '#fff' : C.muted, border: `1px solid ${C.border}` }}>{l}</button>
        ))}
      </div>
      <div className="flex flex-wrap gap-3 items-end mt-3">
        <Field label="Fecha"><input type="date" style={inputStyle} value={fecha} onChange={(e) => setFecha(e.target.value)} /></Field>
        <Field label="Monto ($)"><input type="number" style={{ ...inputStyle, width: 110 }} value={monto} onChange={(e) => setMonto(e.target.value)} /></Field>
        {tipo === 'combustible' && <Field label="Km actual"><input type="number" style={{ ...inputStyle, width: 100 }} value={kmActual} onChange={(e) => setKmActual(e.target.value)} /></Field>}
        {tipo === 'mantenimiento' && (
          <Field label="Tipo de reparación">
            <select style={inputStyle} value={subtipo} onChange={(e) => setSubtipo(e.target.value)}>
              <option value="ordinaria">Ordinaria</option>
              <option value="extraordinaria">Extraordinaria</option>
            </select>
          </Field>
        )}
        <Field label="Detalle"><input style={{ ...inputStyle, width: 220 }} value={detalle} onChange={(e) => setDetalle(e.target.value)} placeholder={tipo === 'multa' ? 'Motivo de la multa' : tipo === 'mantenimiento' ? 'Descripción de la reparación' : 'Notas (opcional)'} /></Field>
      </div>
      <button onClick={() => { if (monto !== '') onSubmit(tipo, { fecha, monto, kmActual: kmActual || null, subtipo: tipo === 'mantenimiento' ? subtipo : null, detalle }); }} className="mt-4 px-4 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.cyanDark }}>
        Guardar registro
      </button>
    </div>
  );
}

function SquadraForm({ techs, onSubmit }) {
  const [nombre, setNombre] = useState('');
  const [miembros, setMiembros] = useState([]);
  const toggleMiembro = (id) => setMiembros((p) => p.includes(id) ? p.filter((x) => x !== id) : [...p, id]);
  return (
    <div className="rounded-xl p-4 mt-4" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
      <Field label="Nombre de la squadra"><input style={{ ...inputStyle, width: 240 }} value={nombre} onChange={(e) => setNombre(e.target.value)} placeholder="Ej. Squadra Norte" /></Field>
      <div className="mt-3">
        <div style={{ fontSize: 12, fontWeight: 600, color: C.slate, marginBottom: 6 }}>Integrantes</div>
        <div className="flex flex-wrap gap-2">
          {techs.map((t) => (
            <label key={t.id} className="flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg" style={{ backgroundColor: miembros.includes(t.id) ? '#E4F8F1' : '#F6F8FA', fontSize: 12.5 }}>
              <input type="checkbox" checked={miembros.includes(t.id)} onChange={() => toggleMiembro(t.id)} style={{ width: 13, height: 13, accentColor: C.cyanDark }} />
              {t.nombre}
            </label>
          ))}
        </div>
      </div>
      <button onClick={() => { if (nombre.trim() && miembros.length > 0) onSubmit({ nombre: nombre.trim(), miembros }); }} className="mt-4 px-4 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.cyanDark }}>
        Crear squadra
      </button>
    </div>
  );
}

function EquipoComodatoForm({ onSubmit }) {
  const [tipo, setTipo] = useState('ONT');
  const [modelo, setModelo] = useState('');
  const [serial, setSerial] = useState('');
  const [sector, setSector] = useState('');
  return (
    <div className="rounded-xl p-4 mt-4 flex flex-wrap gap-3 items-end" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
      <Field label="Tipo">
        <select style={{ ...inputStyle, width: 120 }} value={tipo} onChange={(e) => setTipo(e.target.value)}>
          <option value="ONT">ONT</option>
          <option value="CPE">CPE (módem)</option>
        </select>
      </Field>
      <Field label="Modelo"><input style={{ ...inputStyle, width: 180 }} value={modelo} onChange={(e) => setModelo(e.target.value)} /></Field>
      <Field label="Número de serie"><input style={{ ...inputStyle, width: 180, fontFamily: 'JetBrains Mono, monospace' }} value={serial} onChange={(e) => setSerial(e.target.value)} /></Field>
      <Field label="Sector (opcional)"><input style={{ ...inputStyle, width: 150 }} value={sector} onChange={(e) => setSector(e.target.value)} placeholder="Zona / barrio" /></Field>
      <button onClick={() => { if (modelo.trim() && serial.trim()) onSubmit({ tipo, modelo: modelo.trim(), serial: serial.trim(), sector: sector.trim() }); }} className="px-4 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.cyanDark }}>
        Registrar equipo
      </button>
    </div>
  );
}

function VentaForm({ clients, materials, onSubmit }) {
  const [clienteId, setClienteId] = useState(clients[0]?.id || '');
  const [fecha, setFecha] = useState(new Date().toISOString().slice(0, 10));
  const [lineas, setLineas] = useState([]);
  const [selMat, setSelMat] = useState(materials[0]?.id || '');
  const [selQty, setSelQty] = useState('');
  const [selPrecio, setSelPrecio] = useState('');
  const agregarLinea = () => {
    const qty = Number(selQty), precio = Number(selPrecio);
    if (!selMat || qty <= 0 || precio <= 0) return;
    setLineas((p) => [...p, { materialId: selMat, cantidad: qty, precioUnitario: precio }]);
    setSelQty(''); setSelPrecio('');
  };
  const total = lineas.reduce((s, l) => s + l.cantidad * l.precioUnitario, 0);
  return (
    <div className="rounded-xl p-4 mt-4" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
      <div className="flex flex-wrap gap-3 items-end">
        <Field label="Cliente">
          <select style={{ ...inputStyle, width: 200 }} value={clienteId} onChange={(e) => setClienteId(e.target.value)}>
            {clients.map((c) => <option key={c.id} value={c.id}>{c.nombre}</option>)}
          </select>
        </Field>
        <Field label="Fecha"><input type="date" style={inputStyle} value={fecha} onChange={(e) => setFecha(e.target.value)} /></Field>
      </div>
      <div className="flex flex-wrap gap-3 items-end mt-3 pt-3" style={{ borderTop: `1px dashed ${C.border}` }}>
        <Field label="Producto">
          <select style={{ ...inputStyle, width: 220 }} value={selMat} onChange={(e) => setSelMat(e.target.value)}>
            {materials.map((m) => <option key={m.id} value={m.id}>{m.nombre} (stock: {m.stock})</option>)}
          </select>
        </Field>
        <Field label="Cantidad"><input type="number" style={{ ...inputStyle, width: 90 }} value={selQty} onChange={(e) => setSelQty(e.target.value)} /></Field>
        <Field label="Precio unitario ($)"><input type="number" style={{ ...inputStyle, width: 120 }} value={selPrecio} onChange={(e) => setSelPrecio(e.target.value)} /></Field>
        <button onClick={agregarLinea} className="px-3 py-2 rounded-lg text-sm font-semibold" style={{ backgroundColor: '#EEF1F4', color: C.slate }}>+ Agregar línea</button>
      </div>
      {lineas.length > 0 && (
        <div className="mt-3">
          {lineas.map((l, i) => {
            const m = materials.find((x) => x.id === l.materialId);
            return (
              <div key={i} className="flex items-center justify-between py-1.5 px-2 rounded-lg mt-1" style={{ backgroundColor: '#F6F8FA', fontSize: 12.5 }}>
                <span>{l.cantidad} × {m?.nombre || l.materialId} @ {fmt(l.precioUnitario)}</span>
                <span className="flex items-center gap-3">
                  <span style={{ fontFamily: 'JetBrains Mono, monospace', fontWeight: 600 }}>{fmt(l.cantidad * l.precioUnitario)}</span>
                  <button onClick={() => setLineas((p) => p.filter((_, j) => j !== i))}><X size={13} color={C.coral} /></button>
                </span>
              </div>
            );
          })}
          <div className="flex justify-end mt-2" style={{ fontFamily: 'JetBrains Mono, monospace', fontWeight: 700, fontSize: 15, color: C.slate }}>Total: {fmt(total)}</div>
        </div>
      )}
      <button onClick={() => { if (clienteId && lineas.length > 0) onSubmit({ clienteId, fecha, lineas }); }} disabled={lineas.length === 0}
        className="mt-4 px-4 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: lineas.length === 0 ? C.muted : C.cyanDark }}>
        Registrar venta
      </button>
    </div>
  );
}

function EmpresaForm({ onSubmit }) {
  const [nombre, setNombre] = useState('');
  const [giro, setGiro] = useState(GIROS_NEGOCIO[0]);
  const [eslogan, setEslogan] = useState('');
  return (
    <div className="rounded-xl p-4 mt-4 flex flex-wrap gap-3 items-end" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
      <Field label="Nombre / Razón social"><input style={{ ...inputStyle, width: 240 }} value={nombre} onChange={(e) => setNombre(e.target.value)} placeholder="Ej. Constructora XYZ SA" /></Field>
      <Field label="Giro de negocio">
        <select style={{ ...inputStyle, width: 220 }} value={giro} onChange={(e) => setGiro(e.target.value)}>
          {GIROS_NEGOCIO.map((g) => <option key={g} value={g}>{g}</option>)}
        </select>
      </Field>
      <Field label="Eslogan (opcional)"><input style={{ ...inputStyle, width: 220 }} value={eslogan} onChange={(e) => setEslogan(e.target.value)} /></Field>
      <button onClick={() => { if (nombre.trim()) onSubmit({ nombre: nombre.trim(), giro, eslogan: eslogan.trim() }); }} className="px-4 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.cyanDark }}>
        Crear empresa
      </button>
    </div>
  );
}

function FacturaManualForm({ clients, ingresoCategorias, onSubmit }) {
  const [clienteId, setClienteId] = useState(clients[0]?.id || '');
  const [concepto, setConcepto] = useState('');
  const [categoria, setCategoria] = useState(ingresoCategorias[0]?.nombre || '');
  const [fecha, setFecha] = useState(new Date().toISOString().slice(0, 10));
  const [monto, setMonto] = useState('');
  return (
    <div className="rounded-xl p-4 mt-4 flex flex-wrap gap-3 items-end" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
      <Field label="Cliente">
        <select style={{ ...inputStyle, width: 180 }} value={clienteId} onChange={(e) => setClienteId(e.target.value)}>
          {clients.map((c) => <option key={c.id} value={c.id}>{c.nombre}</option>)}
        </select>
      </Field>
      <Field label="Concepto"><input style={{ ...inputStyle, width: 220 }} value={concepto} onChange={(e) => setConcepto(e.target.value)} placeholder="Ej. Venta de cable sobrante" /></Field>
      <Field label="Categoría de ingreso">
        <select style={{ ...inputStyle, width: 200 }} value={categoria} onChange={(e) => setCategoria(e.target.value)}>
          {ingresoCategorias.map((c) => <option key={c.nombre} value={c.nombre}>{c.nombre}</option>)}
        </select>
      </Field>
      <Field label="Fecha"><input type="date" style={inputStyle} value={fecha} onChange={(e) => setFecha(e.target.value)} /></Field>
      <Field label="Monto ($)"><input type="number" style={{ ...inputStyle, width: 120 }} value={monto} onChange={(e) => setMonto(e.target.value)} /></Field>
      <button onClick={() => { if (concepto.trim() && monto !== '' && clienteId) onSubmit({ clienteId, concepto: concepto.trim(), categoria, fecha, monto }); }} className="px-4 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.cyanDark }}>
        Registrar factura
      </button>
    </div>
  );
}

function AnticipoForm({ empleados, onSubmit }) {
  const [empleadoId, setEmpleadoId] = useState(empleados[0]?.id || '');
  const [fecha, setFecha] = useState(new Date().toISOString().slice(0, 10));
  const [monto, setMonto] = useState('');
  const [motivo, setMotivo] = useState('');
  return (
    <div className="rounded-xl p-4 mt-4 flex flex-wrap gap-3 items-end" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
      <Field label="Empleado"><select style={{ ...inputStyle, width: 200 }} value={empleadoId} onChange={(e) => setEmpleadoId(e.target.value)}>{empleados.map((e) => <option key={e.id} value={e.id}>{e.nombre}</option>)}</select></Field>
      <Field label="Fecha"><input type="date" style={inputStyle} value={fecha} onChange={(e) => setFecha(e.target.value)} /></Field>
      <Field label="Monto ($)"><input type="number" style={{ ...inputStyle, width: 120 }} value={monto} onChange={(e) => setMonto(e.target.value)} /></Field>
      <Field label="Motivo (opcional)"><input style={{ ...inputStyle, width: 200 }} value={motivo} onChange={(e) => setMotivo(e.target.value)} /></Field>
      <button onClick={() => { if (empleadoId && monto !== '') onSubmit({ empleadoId, fecha, monto, motivo }); }} className="px-4 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.cyanDark }}>Registrar anticipo</button>
      <div style={{ fontSize: 11, color: C.muted, width: '100%' }}>Se descontará completo en la siguiente nómina de este empleado.</div>
    </div>
  );
}

function PrestamoForm({ empleados, onSubmit }) {
  const [empleadoId, setEmpleadoId] = useState(empleados[0]?.id || '');
  const [fecha, setFecha] = useState(new Date().toISOString().slice(0, 10));
  const [montoTotal, setMontoTotal] = useState('');
  const [cuotas, setCuotas] = useState('3');
  const [motivo, setMotivo] = useState('');
  return (
    <div className="rounded-xl p-4 mt-4 flex flex-wrap gap-3 items-end" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
      <Field label="Empleado"><select style={{ ...inputStyle, width: 200 }} value={empleadoId} onChange={(e) => setEmpleadoId(e.target.value)}>{empleados.map((e) => <option key={e.id} value={e.id}>{e.nombre}</option>)}</select></Field>
      <Field label="Fecha"><input type="date" style={inputStyle} value={fecha} onChange={(e) => setFecha(e.target.value)} /></Field>
      <Field label="Monto total ($)"><input type="number" style={{ ...inputStyle, width: 120 }} value={montoTotal} onChange={(e) => setMontoTotal(e.target.value)} /></Field>
      <Field label="Número de cuotas"><input type="number" min="1" style={{ ...inputStyle, width: 100 }} value={cuotas} onChange={(e) => setCuotas(e.target.value)} /></Field>
      <Field label="Motivo (opcional)"><input style={{ ...inputStyle, width: 200 }} value={motivo} onChange={(e) => setMotivo(e.target.value)} /></Field>
      {montoTotal !== '' && Number(cuotas) > 0 && <div style={{ fontSize: 11.5, color: C.muted, width: '100%' }}>Cuota estimada: {fmt(Number(montoTotal) / Number(cuotas))} por nómina, hasta liquidarse.</div>}
      <button onClick={() => { if (empleadoId && montoTotal !== '' && Number(cuotas) > 0) onSubmit({ empleadoId, fecha, montoTotal, cuotas, motivo }); }} className="px-4 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.cyanDark }}>Registrar préstamo</button>
    </div>
  );
}

// Organigrama simple por niveles, construido a partir del campo reportaA de cada empleado
function Organigrama({ empleados, updateEmpleado }) {
  const raiz = empleados.filter((e) => !e.reportaA || !empleados.some((o) => o.id === e.reportaA));
  const hijosDe = (id) => empleados.filter((e) => e.reportaA === id);
  const Nodo = ({ emp, nivel }) => {
    const hijos = hijosDe(emp.id);
    return (
      <div style={{ marginLeft: nivel * 28, marginTop: 8 }}>
        <div className="inline-flex items-center gap-2 px-3 py-2 rounded-xl" style={{ backgroundColor: emp.rol === 'Técnico' ? '#E4F8F1' : '#EEF1F4', border: `1px solid ${C.border}`, opacity: emp.activo === false ? 0.5 : 1 }}>
          <div style={{ width: 26, height: 26, borderRadius: '50%', backgroundColor: C.navy, color: '#fff', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 11, fontWeight: 700 }}>
            {emp.nombre.split(' ').map((p) => p[0]).slice(0, 2).join('')}
          </div>
          <div>
            <div style={{ fontSize: 12.5, fontWeight: 600, color: C.slate }}>{emp.nombre}</div>
            <div style={{ fontSize: 10.5, color: C.muted }}>{emp.puesto || emp.rol}</div>
          </div>
        </div>
        {hijos.map((h) => <Nodo key={h.id} emp={h} nivel={nivel + 1} />)}
      </div>
    );
  };
  if (empleados.length === 0) return <div style={{ color: C.muted, fontSize: 13, marginTop: 12 }}>Sin empleados registrados.</div>;
  return (
    <div className="rounded-xl p-5 mt-4" style={{ backgroundColor: C.card, border: `1px solid ${C.border}`, overflowX: 'auto' }}>
      {raiz.map((e) => <Nodo key={e.id} emp={e} nivel={0} />)}
    </div>
  );
}

function CentroForm({ onSubmit }) {
  const [id, setId] = useState('');
  const [nombre, setNombre] = useState('');
  return (
    <div className="rounded-xl p-4 mt-3 flex flex-wrap gap-3 items-end" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
      <Field label="Código"><input style={{ ...inputStyle, width: 120 }} value={id} onChange={(e) => setId(e.target.value)} placeholder="Ej. CC-VEN" /></Field>
      <Field label="Nombre del centro"><input style={{ ...inputStyle, width: 240 }} value={nombre} onChange={(e) => setNombre(e.target.value)} placeholder="Ej. Ventas" /></Field>
      <button onClick={() => { if (id.trim() && nombre.trim()) onSubmit({ id: id.trim(), nombre: nombre.trim() }); }} className="px-4 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.cyanDark }}>
        Guardar centro
      </button>
    </div>
  );
}

function CategoriaForm({ cuentas, onSubmit }) {
  const [nombre, setNombre] = useState('');
  const [cuenta, setCuenta] = useState(cuentas[0]?.codigo || '');
  return (
    <div className="rounded-xl p-4 mt-3 flex flex-wrap gap-3 items-end" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
      <Field label="Nombre de la categoría"><input style={{ ...inputStyle, width: 260 }} value={nombre} onChange={(e) => setNombre(e.target.value)} /></Field>
      <Field label="Cuenta contable">
        <select style={{ ...inputStyle, width: 260 }} value={cuenta} onChange={(e) => setCuenta(e.target.value)}>
          {cuentas.map((c) => <option key={c.codigo} value={c.codigo}>{c.codigo} — {c.nombre}</option>)}
        </select>
      </Field>
      <button onClick={() => { if (nombre.trim() && cuenta) onSubmit({ nombre: nombre.trim(), cuenta }); }} className="px-4 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.cyanDark }}>
        Guardar categoría
      </button>
    </div>
  );
}

function CuentaForm({ existentes, onSubmit }) {
  const [codigo, setCodigo] = useState('');
  const [nombre, setNombre] = useState('');
  const [tipo, setTipo] = useState('Gasto');
  const [grupo, setGrupo] = useState('');
  const [subtipo, setSubtipo] = useState('');
  const yaExiste = existentes.some((c) => c.codigo === codigo.trim());
  return (
    <div className="rounded-xl p-4 mt-4 flex flex-wrap gap-3 items-end" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
      <Field label="Código"><input style={{ ...inputStyle, width: 100 }} value={codigo} onChange={(e) => setCodigo(e.target.value)} placeholder="Ej. 5320" /></Field>
      <Field label="Nombre de la cuenta"><input style={{ ...inputStyle, width: 260 }} value={nombre} onChange={(e) => setNombre(e.target.value)} /></Field>
      <Field label="Tipo">
        <select style={inputStyle} value={tipo} onChange={(e) => { setTipo(e.target.value); setGrupo(''); setSubtipo(''); }}>
          {['Activo', 'Pasivo', 'Capital', 'Ingreso', 'Costo', 'Gasto'].map((t) => <option key={t}>{t}</option>)}
        </select>
      </Field>
      {['Activo', 'Pasivo'].includes(tipo) && (
        <Field label="Clasificación (Balance General)">
          <select style={{ ...inputStyle, width: 200 }} value={subtipo} onChange={(e) => setSubtipo(e.target.value)}>
            <option value="">— Sin clasificar —</option>
            {BG_SUBTIPOS.filter((s) => s.tipo === tipo).map((s) => <option key={s.key} value={s.key}>{s.label}</option>)}
          </select>
        </Field>
      )}
      {['Ingreso', 'Costo', 'Gasto'].includes(tipo) && (
        <Field label="Grupo (Estado de Resultados)">
          <select style={{ ...inputStyle, width: 220 }} value={grupo} onChange={(e) => setGrupo(e.target.value)}>
            <option value="">— Sin agrupar —</option>
            {ER_GRUPOS.map((g) => <option key={g.key} value={g.key}>{g.label}</option>)}
          </select>
        </Field>
      )}
      <button onClick={() => { if (codigo.trim() && nombre.trim() && !yaExiste) onSubmit({ codigo: codigo.trim(), nombre: nombre.trim(), tipo, grupo, subtipo }); }} className="px-4 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.cyanDark }}>
        Guardar cuenta
      </button>
      {yaExiste && <div style={{ fontSize: 11.5, color: C.coral, width: '100%' }}>Ya existe una cuenta con ese código.</div>}
    </div>
  );
}

function AsientoForm({ cuentas, centros, initial, onSubmit }) {
  const [fecha, setFecha] = useState(initial?.fecha || new Date().toISOString().slice(0, 10));
  const [glosa, setGlosa] = useState(initial?.glosa || '');
  const [cuentaDebe, setCuentaDebe] = useState(initial?.cuentaDebe || cuentas[0].codigo);
  const [montoDebe, setMontoDebe] = useState(initial ? String(initial.montoDebe) : '');
  const [cuentaHaber, setCuentaHaber] = useState(initial?.cuentaHaber || cuentas[1].codigo);
  const [montoHaber, setMontoHaber] = useState(initial ? String(initial.montoHaber) : '');
  const [centroCostoId, setCentroCostoId] = useState(initial?.centroCostoId || '');
  return (
    <div>
      <div className="flex gap-3 mb-3">
        <Field label="Fecha"><input type="date" style={{ ...inputStyle, width: 140 }} value={fecha} onChange={(e) => setFecha(e.target.value)} /></Field>
        <div style={{ flex: 1 }}><Field label="Concepto"><input style={{ ...inputStyle, width: '100%' }} value={glosa} onChange={(e) => setGlosa(e.target.value)} /></Field></div>
      </div>
      <div className="mb-3"><Field label="Centro de costo (opcional)">
        <select style={{ ...inputStyle, width: '100%' }} value={centroCostoId} onChange={(e) => setCentroCostoId(e.target.value)}>
          <option value="">Sin centro</option>
          {centros.map((c) => <option key={c.id} value={c.id}>{c.nombre}</option>)}
        </select>
      </Field></div>

      <div className="rounded-lg overflow-hidden" style={{ border: `1px solid ${C.border}` }}>
        <table style={{ width: '100%' }}>
          <thead><tr style={{ backgroundColor: '#F7F8FA' }}><th style={{ fontSize: 10.5 }}>Cuenta</th><th style={{ fontSize: 10.5, width: 110 }}>Debe</th><th style={{ fontSize: 10.5, width: 110 }}>Haber</th></tr></thead>
          <tbody>
            <tr>
              <td>
                <select style={{ ...inputStyle, width: '100%', fontSize: 12 }} value={cuentaDebe} onChange={(e) => setCuentaDebe(e.target.value)}>
                  {cuentas.map((c) => <option key={c.codigo} value={c.codigo}>{c.codigo} — {c.nombre}</option>)}
                </select>
              </td>
              <td><input type="number" style={{ ...inputStyle, width: '100%', fontFamily: 'JetBrains Mono, monospace' }} value={montoDebe} onChange={(e) => { setMontoDebe(e.target.value); setMontoHaber(e.target.value); }} /></td>
              <td></td>
            </tr>
            <tr>
              <td style={{ borderTop: `1px solid ${C.border}` }}>
                <select style={{ ...inputStyle, width: '100%', fontSize: 12, fontStyle: 'italic' }} value={cuentaHaber} onChange={(e) => setCuentaHaber(e.target.value)}>
                  {cuentas.map((c) => <option key={c.codigo} value={c.codigo}>{c.codigo} — {c.nombre}</option>)}
                </select>
              </td>
              <td style={{ borderTop: `1px solid ${C.border}` }}></td>
              <td style={{ borderTop: `1px solid ${C.border}` }}><input type="number" style={{ ...inputStyle, width: '100%', fontFamily: 'JetBrains Mono, monospace' }} value={montoHaber} onChange={(e) => setMontoHaber(e.target.value)} /></td>
            </tr>
          </tbody>
        </table>
      </div>
      <div style={{ fontSize: 10.5, color: C.muted, marginTop: 6 }}>Partida doble: el monto de Debe y Haber se copia igual; ajústalo si el asiento lo requiere.</div>

      <div className="flex gap-2 mt-4">
        <button onClick={() => { if (glosa && montoDebe !== '') onSubmit({ fecha, glosa, cuentaDebe, montoDebe, cuentaHaber, montoHaber, centroCostoId: centroCostoId || null }); }}
          className="px-4 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.cyanDark }}>{initial ? 'Guardar cambios' : 'Registrar asiento'}</button>
      </div>
    </div>
  );
}

function MovimientoForm({ gastoCategorias, ingresoCategorias, onSubmit }) {
  const [tipo, setTipo] = useState('gasto');
  const [categoria, setCategoria] = useState(gastoCategorias[0]?.nombre || '');
  const [descripcion, setDescripcion] = useState('');
  const [fecha, setFecha] = useState(new Date().toISOString().slice(0, 10));
  const [monto, setMonto] = useState('');
  const cats = tipo === 'gasto' ? gastoCategorias : ingresoCategorias;
  return (
    <div className="rounded-xl p-4 mt-4 flex flex-wrap gap-3 items-end" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
      <Field label="Tipo">
        <select style={{ ...inputStyle, width: 120 }} value={tipo} onChange={(e) => { setTipo(e.target.value); setCategoria(e.target.value === 'gasto' ? (gastoCategorias[0]?.nombre || '') : (ingresoCategorias[0]?.nombre || '')); }}>
          <option value="gasto">Gasto</option>
          <option value="ingreso">Ingreso</option>
        </select>
      </Field>
      <Field label="Categoría">
        <select style={{ ...inputStyle, width: 230 }} value={categoria} onChange={(e) => setCategoria(e.target.value)}>
          {cats.map((c) => <option key={c.nombre} value={c.nombre}>{c.nombre}</option>)}
        </select>
      </Field>
      <Field label="Descripción"><input style={{ ...inputStyle, width: 220 }} value={descripcion} onChange={(e) => setDescripcion(e.target.value)} /></Field>
      <Field label="Fecha"><input type="date" style={inputStyle} value={fecha} onChange={(e) => setFecha(e.target.value)} /></Field>
      <Field label="Monto ($)"><input type="number" style={{ ...inputStyle, width: 120 }} value={monto} onChange={(e) => setMonto(e.target.value)} /></Field>
      <button onClick={() => { if (descripcion && monto !== '' && categoria) onSubmit({ tipo, categoria, descripcion, fecha, monto }); }} className="px-4 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.cyanDark }}>
        Guardar
      </button>
    </div>
  );
}

function SimpleForm({ fields, onSubmit }) {
  const [data, setData] = useState(Object.fromEntries(fields.map((f) => [f.key, f.type === 'select' ? (f.options?.[0] || '') : ''])));
  return (
    <div className="rounded-xl p-4 mt-4 flex flex-wrap gap-3 items-end" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
      {fields.map((f) => (
        <Field key={f.key} label={f.label}>
          {f.type === 'select' ? (
            <select value={data[f.key]} onChange={(e) => setData({ ...data, [f.key]: e.target.value })} style={{ ...inputStyle, width: 190 }}>
              {(f.options || []).map((op) => <option key={op} value={op}>{op}</option>)}
            </select>
          ) : (
            <input type={f.type || 'text'} value={data[f.key]} onChange={(e) => setData({ ...data, [f.key]: e.target.value })} style={{ ...inputStyle, width: 190 }} />
          )}
        </Field>
      ))}
      <button onClick={() => { if (fields.every((f) => f.optional || f.type === 'number' || data[f.key] !== '')) onSubmit(data); }} className="px-4 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.cyanDark }}>
        Guardar
      </button>
    </div>
  );
}

function ProjectFormFields({ clients, onSubmit }) {
  const [nombre, setNombre] = useState('');
  const [clienteId, setClienteId] = useState(clients[0]?.id || '');
  const [fechaInicio, setFechaInicio] = useState(new Date().toISOString().slice(0, 10));
  const [presupuesto, setPresupuesto] = useState('');
  return (
    <>
      <Field label="Nombre del proyecto"><input style={{ ...inputStyle, width: 260 }} value={nombre} onChange={(e) => setNombre(e.target.value)} /></Field>
      <Field label="Cliente"><select style={{ ...inputStyle, width: 200 }} value={clienteId} onChange={(e) => setClienteId(e.target.value)}>{clients.map((c) => <option key={c.id} value={c.id}>{c.nombre}</option>)}</select></Field>
      <Field label="Fecha de inicio"><input type="date" style={inputStyle} value={fechaInicio} onChange={(e) => setFechaInicio(e.target.value)} /></Field>
      <Field label="Presupuesto ($)"><input type="number" style={{ ...inputStyle, width: 140 }} value={presupuesto} onChange={(e) => setPresupuesto(e.target.value)} /></Field>
      <button onClick={() => { if (nombre && presupuesto !== '') onSubmit({ nombre, clienteId, fechaInicio, presupuesto: Number(presupuesto) || 0 }); }} className="px-4 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.cyanDark }}>
        Guardar
      </button>
    </>
  );
}

function EntradaForm({ materials, defaultMaterial, centros, onSubmit }) {
  const [materialId, setMaterialId] = useState(defaultMaterial);
  const [cantidad, setCantidad] = useState('');
  const [costoUnitario, setCostoUnitario] = useState('');
  const [fecha, setFecha] = useState(new Date().toISOString().slice(0, 10));
  const [referencia, setReferencia] = useState('');
  const [centroCostoId, setCentroCostoId] = useState('');
  React.useEffect(() => setMaterialId(defaultMaterial), [defaultMaterial]);
  return (
    <div className="rounded-xl p-4 mt-4 flex flex-wrap gap-3 items-end" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
      <Field label="Material">
        <select style={{ ...inputStyle, width: 220 }} value={materialId} onChange={(e) => setMaterialId(e.target.value)}>
          {materials.map((m) => <option key={m.id} value={m.id}>{m.nombre}</option>)}
        </select>
      </Field>
      <Field label="Cantidad"><input type="number" style={{ ...inputStyle, width: 100 }} value={cantidad} onChange={(e) => setCantidad(e.target.value)} /></Field>
      <Field label="Costo unitario ($)"><input type="number" style={{ ...inputStyle, width: 120 }} value={costoUnitario} onChange={(e) => setCostoUnitario(e.target.value)} /></Field>
      <Field label="Fecha"><input type="date" style={inputStyle} value={fecha} onChange={(e) => setFecha(e.target.value)} /></Field>
      <Field label="Centro de costo (opcional)">
        <select style={{ ...inputStyle, width: 180 }} value={centroCostoId} onChange={(e) => setCentroCostoId(e.target.value)}>
          <option value="">Sin centro</option>
          {centros.map((c) => <option key={c.id} value={c.id}>{c.nombre}</option>)}
        </select>
      </Field>
      <Field label="Referencia (proveedor / folio)"><input style={{ ...inputStyle, width: 200 }} value={referencia} onChange={(e) => setReferencia(e.target.value)} /></Field>
      <button onClick={() => { if (materialId && cantidad !== '' && costoUnitario !== '') { onSubmit({ materialId, cantidad, costoUnitario, fecha, referencia, centroCostoId: centroCostoId || null }); setCantidad(''); setCostoUnitario(''); setReferencia(''); } }}
        className="px-4 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.cyanDark }}>
        Registrar entrada
      </button>
    </div>
  );
}

function SalidaForm({ materials, defaultMaterial, centros, orders, onSubmit }) {
  const [materialId, setMaterialId] = useState(defaultMaterial);
  const [cantidad, setCantidad] = useState('');
  const [fecha, setFecha] = useState(new Date().toISOString().slice(0, 10));
  const [referencia, setReferencia] = useState('');
  const [ordenId, setOrdenId] = useState('');
  const [centroCostoId, setCentroCostoId] = useState(centros[0]?.id || '');
  const [aviso, setAviso] = useState('');
  React.useEffect(() => setMaterialId(defaultMaterial), [defaultMaterial]);
  const mat = materials.find((m) => m.id === materialId);
  const ordenSel = orders.find((o) => o.id === ordenId);
  return (
    <div className="rounded-xl p-4 mt-4 flex flex-wrap gap-3 items-end" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
      <Field label="Material">
        <select style={{ ...inputStyle, width: 220 }} value={materialId} onChange={(e) => setMaterialId(e.target.value)}>
          {materials.map((m) => <option key={m.id} value={m.id}>{m.nombre} (stock {m.stock})</option>)}
        </select>
      </Field>
      <Field label="Cantidad"><input type="number" min="1" max={mat?.stock || 0} style={{ ...inputStyle, width: 100 }} value={cantidad} onChange={(e) => setCantidad(e.target.value)} /></Field>
      <Field label="Fecha"><input type="date" style={inputStyle} value={fecha} onChange={(e) => setFecha(e.target.value)} /></Field>
      <Field label="Orden de trabajo (opcional)">
        <select style={{ ...inputStyle, width: 220 }} value={ordenId} onChange={(e) => {
          setOrdenId(e.target.value);
          const o = orders.find((x) => x.id === e.target.value);
          if (o) { setReferencia(`Orden ${o.id}`); if (o.centroCostoId) setCentroCostoId(o.centroCostoId); }
        }}>
          <option value="">— Sin orden (merma / uso interno) —</option>
          {orders.map((o) => <option key={o.id} value={o.id}>{o.id} — {o.direccion || o.tipoServicio}</option>)}
        </select>
      </Field>
      <Field label="Centro de costo">
        <select style={{ ...inputStyle, width: 180 }} value={centroCostoId} onChange={(e) => setCentroCostoId(e.target.value)}>
          {centros.map((c) => <option key={c.id} value={c.id}>{c.nombre}</option>)}
        </select>
      </Field>
      <Field label="Referencia / motivo"><input style={{ ...inputStyle, width: 200 }} value={referencia} onChange={(e) => setReferencia(e.target.value)} placeholder="Ej. merma, uso interno..." /></Field>
      <button onClick={() => {
        if (!materialId) { setAviso('Elige un material.'); return; }
        if (cantidad === '' || Number(cantidad) <= 0) { setAviso('Escribe una cantidad mayor a cero.'); return; }
        if (mat && Number(cantidad) > mat.stock) { setAviso(`No hay suficiente stock — solo hay ${mat.stock}.`); return; }
        setAviso('');
        onSubmit({ materialId, cantidad, fecha, referencia, centroCostoId, ordenSupabaseId: ordenSel?.supabaseId || null });
        setCantidad(''); setReferencia(''); setOrdenId('');
      }} className="px-4 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.coral }}>
        Registrar salida
      </button>
      {aviso && <div style={{ fontSize: 12, color: C.coral, fontWeight: 600, width: '100%' }}>{aviso}</div>}
      <div style={{ fontSize: 11, color: C.muted, width: '100%' }}>Si eliges una orden de trabajo, el material queda ligado a ella (aparecerá en su detalle) y toma su centro de costo automáticamente. Si la dejas en blanco, es para consumos que no vienen de una orden (merma, uso interno, ajustes).</div>
    </div>
  );
}

function OrderForm({ clients, techs, materials, projects, centros, catalogoPartidas, tiposServicio, onSubmit }) {
  const [clienteId, setClienteId] = useState(clients[0]?.id || '');
  const [tecnicoId, setTecnicoId] = useState(techs[0]?.id || '');
  const [proyectoId, setProyectoId] = useState('');
  const [centroCostoId, setCentroCostoId] = useState(centros[0]?.id || '');
  const [tipoServicio, setTipoServicio] = useState('Instalación residencial');
  const [direccion, setDireccion] = useState('');
  const [fecha, setFecha] = useState(new Date().toISOString().slice(0, 10));
  const [costoManoObra, setCostoManoObra] = useState(0);
  const [precioCliente, setPrecioCliente] = useState('');
  const [items, setItems] = useState([]);
  const [selMat, setSelMat] = useState(materials[0]?.id || '');
  const [selQty, setSelQty] = useState(1);
  const [partidas, setPartidas] = useState([]);
  const [selPartida, setSelPartida] = useState(catalogoPartidas[0]?.id || '');
  const [selPartidaPrecio, setSelPartidaPrecio] = useState(catalogoPartidas[0]?.precioRef || 0);

  const addPartida = () => {
    const cat = catalogoPartidas.find((p) => p.id === selPartida);
    if (!cat) return;
    setPartidas((p) => [...p, { partidaId: cat.id, nombre: cat.nombre, precio: Number(selPartidaPrecio) || 0 }]);
  };
  const removePartida = (idx) => setPartidas((p) => p.filter((_, i) => i !== idx));
  const totalPartidas = partidas.reduce((s, p) => s + p.precio, 0);

  const addItem = () => { if (!selMat) return; setItems((p) => [...p, { materialId: selMat, cantidad: Number(selQty) || 1 }]); };
  const removeItem = (idx) => setItems((p) => p.filter((_, i) => i !== idx));

  return (
    <div className="rounded-xl p-5 mt-4" style={{ backgroundColor: C.card, border: `1px solid ${C.border}` }}>
      <div className="grid grid-cols-3 gap-3">
        <Field label="Cliente">
          <select style={inputStyle} value={clienteId} onChange={(e) => { setClienteId(e.target.value); const c = clients.find((x) => x.id === e.target.value); if (c) setDireccion(c.direccion); }}>
            {clients.map((c) => <option key={c.id} value={c.id}>{c.nombre}</option>)}
          </select>
        </Field>
        <Field label="Técnico asignado">
          <select style={inputStyle} value={tecnicoId} onChange={(e) => setTecnicoId(e.target.value)}>
            {techs.map((t) => <option key={t.id} value={t.id}>{t.nombre}</option>)}
          </select>
        </Field>
        <Field label="Proyecto (opcional)">
          <select style={inputStyle} value={proyectoId} onChange={(e) => setProyectoId(e.target.value)}>
            <option value="">— Orden independiente —</option>
            {projects.map((p) => <option key={p.id} value={p.id}>{p.nombre}</option>)}
          </select>
        </Field>
        <Field label="Tipo de servicio">
          <select style={inputStyle} value={tipoServicio} onChange={(e) => setTipoServicio(e.target.value)}>
            {tiposServicio.map((s) => <option key={s}>{s}</option>)}
          </select>
        </Field>
        <Field label="Centro de costo">
          <select style={inputStyle} value={centroCostoId} onChange={(e) => setCentroCostoId(e.target.value)}>
            {centros.map((c) => <option key={c.id} value={c.id}>{c.nombre}</option>)}
          </select>
        </Field>
        <Field label="Dirección"><input style={inputStyle} value={direccion} onChange={(e) => setDireccion(e.target.value)} /></Field>
        <Field label="Fecha"><input type="date" style={inputStyle} value={fecha} onChange={(e) => setFecha(e.target.value)} /></Field>
        <Field label="Precio al cliente ($)"><input type="number" style={inputStyle} value={precioCliente} onChange={(e) => setPrecioCliente(e.target.value)} /></Field>
        <Field label="Costo mano de obra ($)"><input type="number" style={inputStyle} value={costoManoObra} onChange={(e) => setCostoManoObra(e.target.value)} /></Field>
      </div>

      <div className="mt-4 pt-4" style={{ borderTop: `1px solid ${C.border}` }}>
        <div style={{ fontSize: 12.5, fontWeight: 600, color: C.slate, marginBottom: 8 }}>Partidas de trabajo (desglose para facturar al operador)</div>
        <div className="flex gap-2 items-end mb-3">
          <Field label="Partida">
            <select style={{ ...inputStyle, width: 240 }} value={selPartida} onChange={(e) => { setSelPartida(e.target.value); const c = catalogoPartidas.find((p) => p.id === e.target.value); if (c) setSelPartidaPrecio(c.precioRef); }}>
              {catalogoPartidas.map((p) => <option key={p.id} value={p.id}>{p.nombre}</option>)}
            </select>
          </Field>
          <Field label="Precio ($)"><input type="number" style={{ ...inputStyle, width: 100 }} value={selPartidaPrecio} onChange={(e) => setSelPartidaPrecio(e.target.value)} /></Field>
          <button onClick={addPartida} className="px-3 py-1.5 rounded-lg text-sm font-semibold" style={{ backgroundColor: C.bg, color: C.slate, border: `1px solid ${C.border}` }}>Añadir</button>
        </div>
        {partidas.length > 0 && (
          <div className="flex flex-col gap-1.5 mb-2">
            {partidas.map((p, idx) => (
              <div key={idx} className="flex items-center justify-between text-xs px-3 py-1.5 rounded-lg" style={{ backgroundColor: C.bg }}>
                <span>{p.nombre} — {fmt(p.precio)}</span>
                <button onClick={() => removePartida(idx)}><X size={13} color={C.coral} /></button>
              </div>
            ))}
            <div className="text-xs font-semibold" style={{ color: C.cyanDark, textAlign: 'right' }}>Subtotal partidas: {fmt(totalPartidas)} {partidas.length > 0 && Number(precioCliente) !== totalPartidas && '· distinto al precio al cliente capturado arriba'}</div>
          </div>
        )}
        <div style={{ fontSize: 11, color: C.muted }}>El precio al cliente arriba es el que se factura. Si usas partidas, cópialo aquí como referencia del desglose interno.</div>
      </div>

      <div className="mt-4 pt-4" style={{ borderTop: `1px solid ${C.border}` }}>
        <div style={{ fontSize: 12.5, fontWeight: 600, color: C.slate, marginBottom: 8 }}>Materiales a utilizar</div>
        <div className="flex gap-2 items-end mb-3">
          <Field label="Material">
            <select style={{ ...inputStyle, width: 220 }} value={selMat} onChange={(e) => setSelMat(e.target.value)}>
              {materials.map((m) => <option key={m.id} value={m.id}>{m.nombre} — {fmt(m.costoUnitario)}/{m.unidad} (stock {m.stock})</option>)}
            </select>
          </Field>
          <Field label="Cantidad"><input type="number" min="1" style={{ ...inputStyle, width: 80 }} value={selQty} onChange={(e) => setSelQty(e.target.value)} /></Field>
          <button onClick={addItem} className="px-3 py-1.5 rounded-lg text-sm font-semibold" style={{ backgroundColor: C.bg, color: C.slate, border: `1px solid ${C.border}` }}>Añadir</button>
        </div>
        {items.length > 0 && (
          <div className="flex flex-col gap-1.5 mb-2">
            {items.map((it, idx) => {
              const m = materials.find((x) => x.id === it.materialId);
              return (
                <div key={idx} className="flex items-center justify-between text-xs px-3 py-1.5 rounded-lg" style={{ backgroundColor: C.bg }}>
                  <span>{m?.nombre} × {it.cantidad} = {fmt((m?.costoUnitario || 0) * it.cantidad)}</span>
                  <button onClick={() => removeItem(idx)}><X size={13} color={C.coral} /></button>
                </div>
              );
            })}
          </div>
        )}
      </div>

      <button
        onClick={() => onSubmit({ clienteId, tecnicoId, proyectoId: proyectoId || null, centroCostoId, tipoServicio, direccion, fecha, costoManoObra: Number(costoManoObra) || 0, precioCliente: Number(precioCliente) || 0, materialesUsados: items, partidas })}
        className="mt-4 px-4 py-2 rounded-lg text-white text-sm font-semibold" style={{ backgroundColor: C.cyanDark }}>
        Crear orden de trabajo
      </button>
    </div>
  );
}
