-- ============================================================
-- Cambios aditivos — no se borra ni modifica ningún dato existente.
-- Todas las columnas nuevas quedan en null para los registros que
-- ya tenías cargados; vos las vas completando cuando quieras.
-- ============================================================

-- Viajes: código para poder buscar quién hizo un viaje puntual
alter table viajes add column if not exists codigo text;
create index if not exists idx_viajes_codigo on viajes(codigo);

-- Neumáticos: marca y tipo de dibujo (de taco / lisas)
alter table neumaticos add column if not exists marca text;
alter table neumaticos add column if not exists tipo text;

-- Acoplados: vencimiento de válvulas (solo aplica a los que son tipo TOLVA)
alter table acoplados add column if not exists venc_valvula1 date;
alter table acoplados add column if not exists venc_valvula2 date;

-- Choferes: vencimientos y fechas de documentación
alter table choferes add column if not exists vencimiento_dni date;
alter table choferes add column if not exists vencimiento_licencia date;
alter table choferes add column if not exists fecha_curso_defensivo date;
alter table choferes add column if not exists fecha_estudios_medicos date;
alter table choferes add column if not exists fecha_svo date;
alter table choferes add column if not exists fecha_art date;
alter table choferes add column if not exists fecha_eepp date;
