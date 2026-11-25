-- =========================================================
-- Food Cost – Base de datos en Supabase
-- Pegar completo en Supabase > SQL Editor > New query > Run
-- =========================================================

-- 1. Configuración del restaurante (una fila por usuario)
create table if not exists public.configuracion (
  user_id          uuid primary key default auth.uid() references auth.users(id) on delete cascade,
  nombre           text not null default 'Mi restaurante',
  logo             text,                                   -- imagen en base64 (160x160)
  meta_fc          numeric(5,2) not null default 30,
  precios_con_igv  boolean not null default true,
  periodo_ini      date,
  periodo_fin      date,
  updated_at       timestamptz not null default now()
);

-- 2. Insumos
create table if not exists public.insumos (
  id           uuid primary key default gen_random_uuid(),
  user_id      uuid not null default auth.uid() references auth.users(id) on delete cascade,
  codigo       text,
  nombre       text not null,
  categoria    text,
  unidad       text not null check (unidad in ('kg','lt','und')),
  costo        numeric(12,4) not null default 0 check (costo >= 0),
  merma        numeric(5,2)  not null default 0 check (merma >= 0 and merma < 100),
  inv_inicial  numeric(12,3) not null default 0,
  compras      numeric(12,3) not null default 0,
  inv_final    numeric(12,3) not null default 0,
  created_at   timestamptz not null default now()
);

-- 3. Recetas (platos)
create table if not exists public.recetas (
  id                 uuid primary key default gen_random_uuid(),
  user_id            uuid not null default auth.uid() references auth.users(id) on delete cascade,
  codigo             text,
  nombre             text not null,
  categoria          text,
  precio             numeric(10,2) not null default 0 check (precio >= 0),
  unidades_vendidas  numeric(12,2) not null default 0,
  created_at         timestamptz not null default now()
);

-- 4. Ingredientes de cada receta (cantidad por porción)
create table if not exists public.receta_ingredientes (
  id         uuid primary key default gen_random_uuid(),
  user_id    uuid not null default auth.uid() references auth.users(id) on delete cascade,
  receta_id  uuid not null references public.recetas(id) on delete cascade,
  insumo_id  uuid not null references public.insumos(id) on delete cascade,
  cantidad   numeric(12,4) not null default 0 check (cantidad >= 0),
  orden      int not null default 0
);

-- 5. Detalle de ventas cargadas desde Excel (historial)
create table if not exists public.ventas (
  id          bigint generated always as identity primary key,
  user_id     uuid not null default auth.uid() references auth.users(id) on delete cascade,
  lote        text,            -- fecha y hora de la carga
  archivo     text,            -- nombre del Excel
  fecha       date,
  plato       text,
  codigo      text,
  unidades    numeric(12,2) not null default 0,
  precio      numeric(10,2),
  receta_id   uuid references public.recetas(id) on delete set null,
  created_at  timestamptz not null default now()
);

-- Índices
create index if not exists insumos_user_idx  on public.insumos(user_id);
create index if not exists recetas_user_idx  on public.recetas(user_id);
create index if not exists ring_user_idx     on public.receta_ingredientes(user_id);
create index if not exists ring_receta_idx   on public.receta_ingredientes(receta_id);
create index if not exists ring_insumo_idx   on public.receta_ingredientes(insumo_id);
create index if not exists ventas_user_idx   on public.ventas(user_id, fecha);

-- =========================================================
-- Seguridad: cada usuario ve y edita solo sus datos (RLS)
-- =========================================================
alter table public.configuracion       enable row level security;
alter table public.insumos             enable row level security;
alter table public.recetas             enable row level security;
alter table public.receta_ingredientes enable row level security;
alter table public.ventas              enable row level security;

do $$
declare t text;
begin
  foreach t in array array['configuracion','insumos','recetas','receta_ingredientes','ventas'] loop
    execute format('drop policy if exists "solo_dueno" on public.%I', t);
    execute format(
      'create policy "solo_dueno" on public.%I for all to authenticated
         using (user_id = (select auth.uid()))
         with check (user_id = (select auth.uid()))', t);
  end loop;
end $$;

-- =========================================================
-- Vista para reportes: costo por porción y food cost por plato
-- =========================================================
create or replace view public.v_food_cost_platos
with (security_invoker = true) as
select
  r.user_id,
  r.codigo,
  r.nombre                                                         as plato,
  r.precio,
  r.unidades_vendidas,
  round(coalesce(sum(ri.cantidad * i.costo / (1 - i.merma/100)),0), 2) as costo_porcion,
  round(case when r.precio > 0 then
        coalesce(sum(ri.cantidad * i.costo / (1 - i.merma/100)),0)
        / (r.precio / case when c.precios_con_igv then 1.18 else 1 end) * 100 end, 1) as food_cost_pct
from public.recetas r
left join public.receta_ingredientes ri on ri.receta_id = r.id
left join public.insumos i             on i.id = ri.insumo_id
left join public.configuracion c       on c.user_id = r.user_id
group by r.id, c.precios_con_igv;

-- =========================================================
-- Acceso desde la web (Data API) solo para usuarios con sesión
-- =========================================================
grant usage on schema public to authenticated;
grant select, insert, update, delete on
  public.configuracion, public.insumos, public.recetas,
  public.receta_ingredientes, public.ventas to authenticated;
grant usage, select on all sequences in schema public to authenticated;
grant select on public.v_food_cost_platos to authenticated;
revoke all on
  public.configuracion, public.insumos, public.recetas,
  public.receta_ingredientes, public.ventas, public.v_food_cost_platos from anon;

-- =========================================================
-- Recargo al consumo (ej. 13%) para calcular el precio neto
-- =========================================================
alter table public.configuracion add column if not exists recargo_consumo numeric(5,2) not null default 0 check (recargo_consumo >= 0 and recargo_consumo <= 30);
create or replace view public.v_food_cost_platos
with (security_invoker = true) as
select r.user_id, r.codigo, r.nombre as plato, r.precio, r.unidades_vendidas,
  round(coalesce(sum(ri.cantidad * i.costo / (1 - i.merma/100)),0), 2) as costo_porcion,
  round(case when r.precio > 0 then
        coalesce(sum(ri.cantidad * i.costo / (1 - i.merma/100)),0)
        / (r.precio / (1 + case when c.precios_con_igv then 0.18 else 0 end + coalesce(c.recargo_consumo,0)/100)) * 100 end, 1) as food_cost_pct
from public.recetas r
left join public.receta_ingredientes ri on ri.receta_id = r.id
left join public.insumos i             on i.id = ri.insumo_id
left join public.configuracion c       on c.user_id = r.user_id
group by r.id, c.precios_con_igv, c.recargo_consumo;
grant select on public.v_food_cost_platos to authenticated;
revoke all on public.v_food_cost_platos from anon;

-- =========================================================
-- Reportes: ventas del paloteo, mozos por día y compras
-- =========================================================
alter table public.ventas add column if not exists mozo text;
alter table public.ventas add column if not exists grupo text;
alter table public.ventas add column if not exists subgrupo text;
alter table public.ventas add column if not exists canal text not null default 'Salón';
alter table public.ventas add column if not exists neto numeric(12,2) not null default 0;
create index if not exists ventas_user_fecha_idx on public.ventas(user_id, fecha);

create table if not exists public.mozo_dia (
  id          bigint generated always as identity primary key,
  user_id     uuid not null default auth.uid() references auth.users(id) on delete cascade,
  fecha       date not null,
  mozo        text not null,
  comensales  numeric(10,2) not null default 0,
  pedidos     numeric(10,2) not null default 0,
  unique (user_id, fecha, mozo)
);

create table if not exists public.compras (
  id          bigint generated always as identity primary key,
  user_id     uuid not null default auth.uid() references auth.users(id) on delete cascade,
  fecha       date not null,
  familia     text,
  subfamilia  text,
  codigo      text,
  articulo    text not null,
  unidad      text,
  precio      numeric(14,4) not null default 0,
  cantidad    numeric(14,4) not null default 0,
  subtotal    numeric(14,2) not null default 0,
  proveedor   text,
  documento   text,
  archivo     text,
  created_at  timestamptz not null default now()
);
create index if not exists compras_user_fecha_idx on public.compras(user_id, fecha);

alter table public.configuracion add column if not exists alerta_fc numeric(5,2) not null default 35;

alter table public.mozo_dia enable row level security;
alter table public.compras  enable row level security;
create policy "solo_dueno" on public.mozo_dia for all to authenticated
  using (user_id = (select auth.uid())) with check (user_id = (select auth.uid()));
create policy "solo_dueno" on public.compras for all to authenticated
  using (user_id = (select auth.uid())) with check (user_id = (select auth.uid()));

grant select, insert, update, delete on public.mozo_dia, public.compras to authenticated;
grant usage, select on all sequences in schema public to authenticated;
