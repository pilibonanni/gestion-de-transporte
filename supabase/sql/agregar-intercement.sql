-- Aditivo, no toca ningún dato existente.
alter table camiones add column if not exists vencimiento_intercement date;
alter table acoplados add column if not exists vencimiento_intercement date;
