-- Row Level Security for public.brand_mappings
-- Run in Supabase Dashboard > SQL Editor.
--
-- Result:
--   * anon (every visitor, using the key in app.js) can only READ mappings
--   * insert / update / delete require a signed-in Supabase Auth user (role "authenticated")
--
-- Note: the web app currently has no Supabase Auth login, so after running this
-- the "บันทึกลง Database" button will be rejected. Edit mappings via the
-- Supabase Table Editor, or add Supabase Auth login to the Admin panel later.

alter table public.brand_mappings enable row level security;

-- Remove any existing policies on this table (e.g. a permissive "allow all" policy)
do $$
declare
    pol record;
begin
    for pol in
        select policyname from pg_policies
        where schemaname = 'public' and tablename = 'brand_mappings'
    loop
        execute format('drop policy %I on public.brand_mappings', pol.policyname);
    end loop;
end $$;

create policy "brand_mappings_read_all"
    on public.brand_mappings for select
    to anon, authenticated
    using (true);

create policy "brand_mappings_insert_authenticated"
    on public.brand_mappings for insert
    to authenticated
    with check (true);

create policy "brand_mappings_update_authenticated"
    on public.brand_mappings for update
    to authenticated
    using (true)
    with check (true);

create policy "brand_mappings_delete_authenticated"
    on public.brand_mappings for delete
    to authenticated
    using (true);

-- Defense in depth: also revoke table-level write privileges from anon
revoke insert, update, delete on public.brand_mappings from anon;
