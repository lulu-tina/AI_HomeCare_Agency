-- CareFlow motorcycle routing draft, based on OSRM's installed car profile.
-- Conservative restrictions and provisional speeds; validate against local routes.
package.path = '/opt/?.lua;' .. package.path
local base = dofile('/opt/car.lua')
local car_setup = base.setup
local car_process_way = base.process_way

base.setup = function()
  local profile = car_setup()
  profile.access_tags_hierarchy = Sequence {'motorcycle', 'motor_vehicle', 'vehicle', 'access'}
  profile.restrictions = Sequence {'motorcycle', 'motor_vehicle', 'vehicle'}
  profile.access_tag_whitelist['motorcycle'] = true
  profile.vehicle_width = 0.8
  profile.vehicle_height = 1.8
  profile.vehicle_length = 2.0
  profile.vehicle_weight = 250
  profile.speeds = Sequence { highway = {
    primary=40, primary_link=25, secondary=35, secondary_link=25,
    tertiary=30, tertiary_link=20, unclassified=25, residential=25,
    living_street=10, service=15
  }}
  profile.default_speed = 15
  return profile
end

base.process_way = function(profile, way, result, relations)
  local highway = way:get_value_by_key('highway')
  local motorcycle = way:get_value_by_key('motorcycle')
  -- Do not assume small motorcycles may use motorway or expressway-like roads.
  -- Trunk classification may exclude some otherwise usable roads: conservative draft.
  if highway == 'motorway' or highway == 'motorway_link' or
     highway == 'trunk' or highway == 'trunk_link' or
     way:get_value_by_key('motorroad') == 'yes' then return end
  if motorcycle == 'no' or motorcycle == 'private' then return end
  if (highway == 'cycleway' or highway == 'footway' or highway == 'path' or
      highway == 'pedestrian' or highway == 'steps') then return end
  return car_process_way(profile, way, result, relations)
end
return base
