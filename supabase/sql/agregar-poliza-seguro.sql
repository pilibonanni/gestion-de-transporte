-- Agrega los campos para guardar la póliza de seguro del camión
alter table camiones add column if not exists poliza_url text;
alter table camiones add column if not exists poliza_nombre text;

-- Crea el bucket de Storage donde se guardan los archivos de las pólizas
-- (privado: solo se accede con un link temporal generado por la app, no público)
insert into storage.buckets (id, name, public)
values ('polizas', 'polizas', false)
on conflict (id) do nothing;

-- Políticas de acceso al bucket: quien tenga permiso de "camiones" puede subir,
-- reemplazar y descargar. Eliminar el archivo queda para el admin.
create policy "ver polizas con permiso"
  on storage.objects for select
  using (bucket_id = 'polizas' and tiene_permiso('camiones'));

create policy "subir polizas con permiso"
  on storage.objects for insert
  with check (bucket_id = 'polizas' and tiene_permiso('camiones'));

create policy "reemplazar polizas con permiso"
  on storage.objects for update
  using (bucket_id = 'polizas' and tiene_permiso('camiones'));

create policy "eliminar polizas solo admin"
  on storage.objects for delete
  using (bucket_id = 'polizas' and es_admin());
