-- 前台登入只回傳前台會使用的欄位；後台 case_file_state 完整資料不變。
-- 此函式已部署至公司 Supabase 專案。
create table if not exists public.case_file_front_access (
  person_id text primary key,
  person_name text not null default '',
  last_entered_at timestamptz not null default now()
);
alter table public.case_file_front_access enable row level security;
drop policy if exists "authenticated can read front access" on public.case_file_front_access;
create policy "authenticated can read front access" on public.case_file_front_access for select to authenticated using (true);
grant select on public.case_file_front_access to authenticated;

create or replace function public.case_file_front_touch(p_national_id text)
returns boolean
language plpgsql
security definer
set search_path = public
as $$
declare
  v_person jsonb;
begin
  select person into v_person
  from public.case_file_state state,
  jsonb_array_elements(coalesce(state.data->'settings'->'personnel', '[]'::jsonb)) person
  where state.id = 'main'
    and upper(regexp_replace(coalesce(person->>'nationalId', ''), '\s+', '', 'g')) = upper(regexp_replace(coalesce(p_national_id, ''), '\s+', '', 'g'))
    and coalesce(person->>'status', '在職') = '在職'
    and length(regexp_replace(coalesce(p_national_id, ''), '\s+', '', 'g')) >= 8
  limit 1;
  if v_person is null then return false; end if;
  insert into public.case_file_front_access(person_id, person_name, last_entered_at)
  values (v_person->>'id', coalesce(v_person->>'name', ''), now())
  on conflict (person_id) do update set person_name = excluded.person_name, last_entered_at = excluded.last_entered_at;
  return true;
end;
$$;
revoke all on function public.case_file_front_touch(text) from public;
grant execute on function public.case_file_front_touch(text) to anon, authenticated;

create or replace function public.case_file_front_login(p_national_id text)
returns jsonb
language plpgsql
security definer
set search_path = public
as $$
declare
  v_data jsonb;
  v_person jsonb;
  v_people jsonb;
  v_front_records jsonb;
begin
  select data into v_data from public.case_file_state where id = 'main';
  if v_data is null then return null; end if;

  select person into v_person
  from jsonb_array_elements(coalesce(v_data->'settings'->'personnel', '[]'::jsonb)) person
  where upper(regexp_replace(coalesce(person->>'nationalId', ''), '\s+', '', 'g')) = upper(regexp_replace(coalesce(p_national_id, ''), '\s+', '', 'g'))
    and coalesce(person->>'status', '在職') = '在職'
    and length(regexp_replace(coalesce(p_national_id, ''), '\s+', '', 'g')) >= 8
  limit 1;
  if v_person is null then return null; end if;

  insert into public.case_file_front_access(person_id, person_name, last_entered_at)
  values (v_person->>'id', coalesce(v_person->>'name', ''), now())
  on conflict (person_id) do update set person_name = excluded.person_name, last_entered_at = excluded.last_entered_at;

  select coalesce(jsonb_agg(person - 'nationalId'), '[]'::jsonb)
  into v_people
  from jsonb_array_elements(coalesce(v_data->'settings'->'personnel', '[]'::jsonb)) person;

  select coalesce(jsonb_agg(filtered), '[]'::jsonb)
  into v_front_records
  from jsonb_array_elements(coalesce(v_data->'records', '[]'::jsonb)) record,
  lateral (
    select coalesce(jsonb_object_agg(e.key, e.value), '{}'::jsonb) as filtered
    from jsonb_each(record) e
    where e.key = any(array[
      'id','propertyNo','entrustStart','entrustEnd','area','caseName','caseNameNote','address','type','contractType','status','archived',
      'price','reducedPrice','direction','builtYear','completionDate','age','titleFloor','currentFloor','floor','layout',
      'indoorPing','buildingPing','landPing','parking','parkingOwnership','parkingType','parkingMethod','parkingNo','managementFee','key','currentState',
      'road','frontage','depth','zoning','coverage','far','developer','reportDate','updateDate','groupViewDate','bookLocationType','bookLocationDate',
      'salesBook','salesBookDate','photoInfo','notes','attentionNotes','additionNotes',
      'platform591','platform591Expiry','platform591None','yes319','yes319None','houseinfor','houseinforNone',
      'windowAd','windowAdNone','led','ledNone','homeWeb','homeWebNone','price5168','price5168Expiry','price5168None',
      'goldExposure','goldExposureExpiry','goldExposureNone','_updateHistory','lastModifiedAt','caseNameNoteModifiedAt','reducedPriceModifiedAt','_dailyAnnotationType',
      '_dailyHighlight','_dailyAnnotation','_monthlyReports','_archiveActionDate','_restoredAt'
    ]::text[])
  ) picked;

  return jsonb_build_object('personId', v_person->>'id', 'records', v_front_records, 'personnel', v_people);
end;
$$;

revoke all on function public.case_file_front_login(text) from public;
grant execute on function public.case_file_front_login(text) to anon, authenticated;
