-- Corrige el error de "foreign key constraint" al eliminar un chofer, camión o
-- acoplado que ya tiene viajes cargados. Los viajes quedan como historial,
-- solo se les saca la referencia al chofer/vehículo eliminado.

-- chofer_id: hay que sacarle el NOT NULL para poder ponerlo en null después
alter table viajes alter column chofer_id drop not null;

alter table viajes drop constraint if exists viajes_chofer_id_fkey;
alter table viajes add constraint viajes_chofer_id_fkey
  foreign key (chofer_id) references choferes(id) on delete set null;

alter table viajes drop constraint if exists viajes_camion_id_fkey;
alter table viajes add constraint viajes_camion_id_fkey
  foreign key (camion_id) references camiones(id) on delete set null;

alter table viajes drop constraint if exists viajes_acoplado_id_fkey;
alter table viajes add constraint viajes_acoplado_id_fkey
  foreign key (acoplado_id) references acoplados(id) on delete set null;
