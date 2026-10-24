-- Corrige el error "stack depth limit exceeded" al eliminar registros.
-- Causa: es_admin() y tiene_permiso() consultan la tabla "perfiles", pero esa
-- consulta dispara de nuevo la política de RLS de "perfiles" (que a su vez
-- llama a es_admin()), generando un bucle infinito.
-- Solución: marcarlas como SECURITY DEFINER para que la consulta interna
-- se salte esa verificación y no vuelva a disparar la política.

create or replace function es_admin()
returns boolean
language sql stable security definer
set search_path = public
as $$
  select coalesce((select role = 'admin' from perfiles where id = auth.uid()), false);
$$;

create or replace function tiene_permiso(modulo text)
returns boolean
language sql stable security definer
set search_path = public
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
