-- =========================================================
-- SISTEMA DE FLOTA — Esquema de base de datos para Supabase
-- =========================================================
-- Pensado para calzar 1:1 con el modelo de datos del prototipo
-- (panel-flota.jsx). Ejecutar en el SQL Editor de Supabase.

create extension if not exists "pgcrypto";

-- ---------------------------------------------------------
-- PERFILES DE USUARIO (vinculado a Supabase Auth)
-- ---------------------------------------------------------
-- Supabase Auth maneja el login (email/password) de forma segura.
-- Esta tabla guarda el rol y los permisos por módulo de cada usuario.
create table perfiles (
  id uuid primary key references auth.users(id) on delete cascade,
  username text unique not null,
  role text not null default 'usuario' check (role in ('admin', 'usuario')),
  permiso_camiones boolean not null default false,
  permiso_acoplados boolean not null default false,
  permiso_choferes boolean not null default false,
  permiso_neumaticos boolean not null default false,
  permiso_rutas boolean not null default false,
  created_at timestamptz not null default now()
);

-- Función helper: ¿el usuario actual es admin?
create or replace function es_admin()
returns boolean
language sql stable
as $$
  select coalesce((select role = 'admin' from perfiles where id = auth.uid()), false);
$$;

-- Función helper: ¿el usuario actual tiene permiso sobre un módulo?
create or replace function tiene_permiso(modulo text)
returns boolean
language sql stable
as $$
  select
    coalesce(
      (select role = 'admin' from perfiles where id = auth.uid()), false
    )
    or coalesce(
      (select
        case modulo
          when 'camiones'   then permiso_camiones
          when 'acoplados'  then permiso_acoplados
          when 'choferes'   then permiso_choferes
          when 'neumaticos' then permiso_neumaticos
          when 'rutas'      then permiso_rutas
          else false
        end
       from perfiles where id = auth.uid()),
      false
    );
$$;

-- ---------------------------------------------------------
-- CAMIONES
-- ---------------------------------------------------------
create table camiones (
  id uuid primary key default gen_random_uuid(),
  patente text unique not null,
  modelo text,
  km integer not null default 0,
  ultimo_service_km integer not null default 0,
  intervalo_service integer not null default 50000,
  vtv_ultima date,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

-- ---------------------------------------------------------
-- ACOPLADOS
-- ---------------------------------------------------------
create table acoplados (
  id uuid primary key default gen_random_uuid(),
  patente text unique not null,
  tipo text,
  km integer not null default 0,
  ultimo_service_km integer not null default 0,
  intervalo_service integer not null default 50000,
  vtv_ultima date,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

-- ---------------------------------------------------------
-- SERVICE (historial, tanto de camión como de acoplado)
-- ---------------------------------------------------------
create table servicios (
  id uuid primary key default gen_random_uuid(),
  camion_id uuid references camiones(id) on delete cascade,
  acoplado_id uuid references acoplados(id) on delete cascade,
  fecha date not null default current_date,
  km integer not null,
  detalle text not null,
  editado boolean not null default false,
  fecha_edicion date,
  created_at timestamptz not null default now(),
  constraint servicio_un_solo_vehiculo check (
    (camion_id is not null and acoplado_id is null) or
    (camion_id is null and acoplado_id is not null)
  )
);

-- ---------------------------------------------------------
-- CHOFERES
-- ---------------------------------------------------------
create table choferes (
  id uuid primary key default gen_random_uuid(),
  nombre text not null,
  licencia text,
  km integer not null default 0,
  camion_id uuid references camiones(id) on delete set null,
  acoplado_id uuid references acoplados(id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

-- ---------------------------------------------------------
-- NEUMÁTICOS
-- ---------------------------------------------------------
create table neumaticos (
  id uuid primary key default gen_random_uuid(),
  etiqueta text not null,
  estado text not null default 'Nuevo' check (estado in ('Nuevo', 'Usado', 'Recapado')),
  km integer not null default 0,
  camion_id uuid references camiones(id) on delete set null,
  acoplado_id uuid references acoplados(id) on delete set null,
  posicion text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

-- Historial cronológico del neumático (altas, recapados, cambios de vehículo/posición)
create table neumatico_historial (
  id uuid primary key default gen_random_uuid(),
  neumatico_id uuid not null references neumaticos(id) on delete cascade,
  fecha date not null default current_date,
  tipo text not null check (tipo in ('alta', 'recapado', 'cambio_vehiculo')),
  detalle text not null,
  created_at timestamptz not null default now()
);

-- ---------------------------------------------------------
-- RUTAS (predefinidas, con km fijo)
-- ---------------------------------------------------------
create table rutas (
  id uuid primary key default gen_random_uuid(),
  nombre text not null,
  km integer not null,
  created_at timestamptz not null default now()
);

-- ---------------------------------------------------------
-- VIAJES (historial de cada viaje realizado)
-- ---------------------------------------------------------
create table viajes (
  id uuid primary key default gen_random_uuid(),
  ruta_id uuid references rutas(id),
  ruta_nombre text not null,   -- copia del nombre al momento del viaje (por si se borra la ruta)
  chofer_id uuid not null references choferes(id),
  camion_id uuid references camiones(id),
  acoplado_id uuid references acoplados(id),
  km integer not null,
  fecha date not null default current_date,
  combo_distinto boolean not null default false,
  created_at timestamptz not null default now()
);

-- ---------------------------------------------------------
-- ÍNDICES
-- ---------------------------------------------------------
create index idx_neumaticos_camion on neumaticos(camion_id);
create index idx_neumaticos_acoplado on neumaticos(acoplado_id);
create index idx_neumatico_historial_neumatico on neumatico_historial(neumatico_id);
create index idx_servicios_camion on servicios(camion_id);
create index idx_servicios_acoplado on servicios(acoplado_id);
create index idx_choferes_camion on choferes(camion_id);
create index idx_choferes_acoplado on choferes(acoplado_id);
create index idx_viajes_chofer on viajes(chofer_id);

-- ---------------------------------------------------------
-- RLS (Row Level Security) — solo usuarios logueados con permiso
-- ---------------------------------------------------------
alter table perfiles enable row level security;
alter table camiones enable row level security;
alter table acoplados enable row level security;
alter table servicios enable row level security;
alter table choferes enable row level security;
alter table neumaticos enable row level security;
alter table neumatico_historial enable row level security;
alter table rutas enable row level security;
alter table viajes enable row level security;

-- Perfiles: cada uno ve el suyo; el admin ve y crea todos
create policy "ver propio perfil" on perfiles for select using (id = auth.uid() or es_admin());
create policy "admin crea perfiles" on perfiles for insert with check (es_admin());
create policy "admin edita perfiles" on perfiles for update using (es_admin());
create policy "admin elimina perfiles" on perfiles for delete using (es_admin());

-- Camiones
create policy "ver camiones con permiso" on camiones for select using (tiene_permiso('camiones'));
create policy "editar camiones con permiso" on camiones for update using (tiene_permiso('camiones'));
create policy "crear camiones solo admin" on camiones for insert with check (es_admin());
create policy "eliminar camiones solo admin" on camiones for delete using (es_admin());

-- Acoplados
create policy "ver acoplados con permiso" on acoplados for select using (tiene_permiso('acoplados'));
create policy "editar acoplados con permiso" on acoplados for update using (tiene_permiso('acoplados'));
create policy "crear acoplados solo admin" on acoplados for insert with check (es_admin());
create policy "eliminar acoplados solo admin" on acoplados for delete using (es_admin());

-- Choferes
create policy "ver choferes con permiso" on choferes for select using (tiene_permiso('choferes'));
create policy "editar choferes con permiso" on choferes for update using (tiene_permiso('choferes'));
create policy "crear choferes solo admin" on choferes for insert with check (es_admin());
create policy "eliminar choferes solo admin" on choferes for delete using (es_admin());

-- Neumáticos
create policy "ver neumaticos con permiso" on neumaticos for select using (tiene_permiso('neumaticos'));
create policy "editar neumaticos con permiso" on neumaticos for update using (tiene_permiso('neumaticos'));
create policy "crear neumaticos solo admin" on neumaticos for insert with check (es_admin());
create policy "eliminar neumaticos solo admin" on neumaticos for delete using (es_admin());

-- Historial de neumático (se crea automáticamente al operar sobre neumáticos)
create policy "ver historial con permiso" on neumatico_historial for select using (tiene_permiso('neumaticos'));
create policy "crear historial con permiso" on neumatico_historial for insert with check (tiene_permiso('neumaticos'));

-- Service (asociado a camión o acoplado, sigue el permiso correspondiente)
create policy "ver servicios con permiso" on servicios for select using (
  (camion_id is not null and tiene_permiso('camiones')) or
  (acoplado_id is not null and tiene_permiso('acoplados'))
);
create policy "crear servicios con permiso" on servicios for insert with check (
  (camion_id is not null and tiene_permiso('camiones')) or
  (acoplado_id is not null and tiene_permiso('acoplados'))
);
create policy "editar servicios con permiso" on servicios for update using (
  (camion_id is not null and tiene_permiso('camiones')) or
  (acoplado_id is not null and tiene_permiso('acoplados'))
);

-- Rutas
create policy "ver rutas con permiso" on rutas for select using (tiene_permiso('rutas'));
create policy "crear rutas solo admin" on rutas for insert with check (es_admin());
create policy "eliminar rutas solo admin" on rutas for delete using (es_admin());

-- Viajes (se cargan desde la ficha del chofer -> permiso de choferes)
create policy "ver viajes con permiso" on viajes for select using (tiene_permiso('choferes'));
create policy "crear viajes con permiso" on viajes for insert with check (tiene_permiso('choferes'));

-- ---------------------------------------------------------
-- TRIGGER: updated_at automático
-- ---------------------------------------------------------
create or replace function set_updated_at()
returns trigger language plpgsql as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

create trigger trg_camiones_updated before update on camiones
  for each row execute function set_updated_at();
create trigger trg_acoplados_updated before update on acoplados
  for each row execute function set_updated_at();
create trigger trg_choferes_updated before update on choferes
  for each row execute function set_updated_at();
create trigger trg_neumaticos_updated before update on neumaticos
  for each row execute function set_updated_at();

-- ---------------------------------------------------------
-- USUARIO ADMIN INICIAL
-- ---------------------------------------------------------
-- 1) Creá el usuario admin desde Authentication > Users en el
--    dashboard de Supabase (con su email y contraseña).
-- 2) Copiá su UUID y corré esto reemplazando los valores:
--
-- insert into perfiles (id, username, role, permiso_camiones, permiso_acoplados, permiso_choferes, permiso_neumaticos, permiso_rutas)
-- values ('UUID-DEL-USUARIO-ADMIN', 'admin', 'admin', true, true, true, true, true);
