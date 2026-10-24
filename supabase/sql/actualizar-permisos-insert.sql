-- Actualiza las políticas de INSERT: ahora cualquier usuario con permiso
-- sobre esa sección puede cargar registros nuevos (antes era solo admin).
-- Eliminar sigue siendo exclusivo del admin (no se toca esa parte).

drop policy if exists "crear camiones solo admin" on camiones;
create policy "crear camiones con permiso" on camiones for insert with check (tiene_permiso('camiones'));

drop policy if exists "crear acoplados solo admin" on acoplados;
create policy "crear acoplados con permiso" on acoplados for insert with check (tiene_permiso('acoplados'));

drop policy if exists "crear choferes solo admin" on choferes;
create policy "crear choferes con permiso" on choferes for insert with check (tiene_permiso('choferes'));

drop policy if exists "crear neumaticos solo admin" on neumaticos;
create policy "crear neumaticos con permiso" on neumaticos for insert with check (tiene_permiso('neumaticos'));

drop policy if exists "crear rutas solo admin" on rutas;
create policy "crear rutas con permiso" on rutas for insert with check (tiene_permiso('rutas'));

-- Los servicios (agregar un registro de service) ya seguían el permiso del vehículo, no admin — sin cambios ahí.
-- El borrado (delete) de todas las tablas sigue exclusivo de admin — sin cambios ahí tampoco.

-- Viajes: antes solo se podían crear y ver, faltaban permisos para editar/eliminar
-- (necesario para poder corregir un viaje mal cargado).
create policy "editar viajes con permiso" on viajes for update using (tiene_permiso('choferes'));
create policy "eliminar viajes con permiso" on viajes for delete using (tiene_permiso('choferes'));
