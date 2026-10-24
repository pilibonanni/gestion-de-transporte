-- Aditivo, no toca ningún dato existente.
alter table viajes add column if not exists remito text;
