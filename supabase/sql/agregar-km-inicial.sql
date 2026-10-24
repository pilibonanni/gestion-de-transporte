-- Agrega una columna "km_inicial" (el kilometraje que tenía el camión/acoplado
-- ANTES de empezar a cargar viajes en el sistema). A partir de ahora, el km
-- "real" de cada camión/acoplado se calcula siempre como:
--   km_inicial + suma de los km de sus viajes cargados
-- en vez de confiar en un contador que se va sumando/restando (y que puede
-- desincronizarse con ediciones/correcciones).

alter table camiones add column if not exists km_inicial integer;
alter table acoplados add column if not exists km_inicial integer;

-- Corrección única: calcula el km_inicial correcto de cada uno ya cargado,
-- a partir de su km actual guardado menos la suma de sus viajes ya cargados.
-- Esto hace que, a partir de ahora, el cálculo dinámico coincida exactamente
-- y nunca más se desincronice.

update camiones c
set km_inicial = c.km - coalesce((select sum(v.km) from viajes v where v.camion_id = c.id), 0)
where km_inicial is null;

update acoplados a
set km_inicial = a.km - coalesce((select sum(v.km) from viajes v where v.acoplado_id = a.id), 0)
where km_inicial is null;
