-- Borra el perfil viejo (apuntaba al usuario anterior, ya eliminado)
delete from perfiles where id = '28bb9d75-011f-4c93-9f93-42ecd55d0983';

-- Crea el perfil admin para el nuevo usuario
insert into perfiles (id, username, role, permiso_camiones, permiso_acoplados, permiso_choferes, permiso_neumaticos, permiso_rutas)
values (
  'f85f963c-0712-4287-9a58-d585c2fbbc22',
  'admin',
  'admin',
  true, true, true, true, true
);
