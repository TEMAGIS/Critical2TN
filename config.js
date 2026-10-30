// ════════════════════════════════════════════════════════════════
//  Critical2TN mobile — settings you can change without touching index.html
//  Keep this file next to index.html. Reload the app after editing.
// ════════════════════════════════════════════════════════════════
window.C2TN_CONFIG = {

  // Members of this ArcGIS group see the "Manage" button (users + map layers).
  // Critical2TN User Management
  manageGroupId: 'fb38085094804e58bc02e35d082ec4bf',

  // Members of this group can edit Critical2TN records and photos (the group that has update
  // rights on the layer). Managers grant/remove it from the Manage screen.
  editGroupId: '5368e53cd08a4e0a97399e4cdc9fbae2',

  // ── MFA ──────────────────────────────────────────────────────
  // false (now): MFA is NOT required. Managers and editors only need their own group.
  //              The user list shows everyone in userListGroups below (Statewide, Staging, Secure MFA).
  // true (later): managers and editors must ALSO be in the Secure MFA group, and the user list
  //              shows Secure MFA members only.
  requireSecureMfa: false,

  secureMfaGroupId: 'dafcbf87b71c48528d197b64785ac10b',      // Critical2TN Secure MFA

  // Whose names appear in the Manage > Users list while requireSecureMfa is false.
  // Add or remove groups freely (counties can be added the same way: { id: '…', label: 'Davidson' }).
  userListGroups: [
    { id: '33fc71dc1fca4fd19fefe8ef5bde9112', label: 'Critical Infrastructure' },  // Critical Infrastructure (approved)
    { id: '25ca696fb06f4e64afb2f0450c066d22', label: 'Staging' },      // Critical2TN Staging (approved, MFA not confirmed yet)
    { id: 'dafcbf87b71c48528d197b64785ac10b', label: 'Secure MFA' },   // Critical2TN Secure MFA
  ],

  // ArcGIS item (type JSON) that stores the shared symbol / colour / zoom choices for the map layers.
  // Leave blank, open Manage > Map layers, and use "Create the settings item"; then paste its ID here.
  mapSettingsItemId: '9d29d3e459bb43f188d4318f9bb16131',

  // ── Linked detail tables (extra survey fields tied to a Critical2TN record by its GlobalID) ──
  // Airports: create the table (Airports/README.md), then paste its URLs below. While `url` is blank the
  // "Airport details" section is switched off.
  //   url      = read-only view/table every Critical2TN user can read
  //   editUrl  = the same table as an update-enabled view/table shared ONLY with the edit group
  detailTables: [
    {
      key: 'airport',
      title: 'Airport details',
      url: 'https://services1.arcgis.com/kILp9lqGUeOhnDbI/arcgis/rest/services/Critical2TN_Airport_Details/FeatureServer/0',
      editUrl: 'https://services1.arcgis.com/kILp9lqGUeOhnDbI/arcgis/rest/services/Critical2TN_Airport_Details/FeatureServer/0',
      linkField: 'critical2tn_guid',
      idField: 'airport_id',
      match: { sector: 'transport', hints: 'airport|airfield|airpark|gliderport|landing strip' },
      groups: [
        {
          "label": "Airport",
          "fields": [
            {
              "field": "airport_id",
              "label": "Airport ID",
              "type": "text",
              "length": 8,
              "required": true
            },
            {
              "field": "sponsor",
              "label": "Legal airport sponsor",
              "type": "text",
              "length": 150
            },
            {
              "field": "distance_direction",
              "label": "Distance / direction from city",
              "type": "text",
              "length": 60
            },
            {
              "field": "tasp_class",
              "label": "TASP classification",
              "type": "select",
              "length": 40,
              "options": [
                "Turf",
                "Community Business",
                "Community Service",
                "Regional Service",
                "Commercial Service"
              ]
            },
            {
              "field": "elevation_ft",
              "label": "Elevation (ft)",
              "type": "number"
            },
            {
              "field": "tema_region",
              "label": "TEMA region",
              "type": "select",
              "length": 20,
              "options": [
                "East",
                "Middle",
                "Southeast",
                "West"
              ]
            },
            {
              "field": "tdot_region",
              "label": "TDOT region",
              "type": "select",
              "length": 4,
              "options": [
                "1",
                "2",
                "3",
                "4"
              ]
            },
            {
              "field": "time_zone",
              "label": "Time zone",
              "type": "select",
              "length": 4,
              "options": [
                "CST",
                "EST"
              ]
            }
          ]
        },
        {
          "label": "Contacts (24 hr)",
          "fields": [
            {
              "field": "primary_contact_name",
              "label": "Primary contact",
              "type": "text",
              "length": 100
            },
            {
              "field": "primary_contact_phone",
              "label": "Primary 24 hr phone",
              "type": "tel",
              "length": 30
            },
            {
              "field": "secondary_contact_name",
              "label": "Secondary contact",
              "type": "text",
              "length": 100
            },
            {
              "field": "secondary_contact_phone",
              "label": "Secondary 24 hr phone",
              "type": "tel",
              "length": 30
            },
            {
              "field": "ema_director_name",
              "label": "Emergency management director",
              "type": "text",
              "length": 100
            },
            {
              "field": "ema_director_phone",
              "label": "EMA director 24 hr phone",
              "type": "tel",
              "length": 30
            }
          ]
        },
        {
          "label": "Runways & capability",
          "fields": [
            {
              "field": "num_runways",
              "label": "Number of runways",
              "type": "number",
              "integer": true
            },
            {
              "field": "runway_lengths",
              "label": "Runway lengths (ft)",
              "type": "text",
              "length": 100
            },
            {
              "field": "runway_width",
              "label": "Runway width (ft)",
              "type": "text",
              "length": 60
            },
            {
              "field": "runway_surface",
              "label": "Runway surface",
              "type": "text",
              "length": 100
            },
            {
              "field": "tang_c17",
              "label": "TANG C-17 landing capable",
              "type": "yesno",
              "length": 3
            },
            {
              "field": "tang_c5",
              "label": "TANG C-5 landing capable",
              "type": "yesno",
              "length": 3
            }
          ]
        },
        {
          "label": "Services",
          "fields": [
            {
              "field": "storage",
              "label": "Storage (hangar / tie-down)",
              "type": "text",
              "length": 150
            },
            {
              "field": "fuel_type",
              "label": "Fuel type",
              "type": "text",
              "length": 100
            },
            {
              "field": "ops_24hr",
              "label": "24 hr operations",
              "type": "yesno",
              "length": 3
            },
            {
              "field": "lodging_onsite",
              "label": "Lodging on site",
              "type": "yesno",
              "length": 3
            }
          ]
        },
        {
          "label": "Survey",
          "fields": [
            {
              "field": "survey_status",
              "label": "Survey status",
              "type": "select",
              "length": 12,
              "options": [
                "Placeholder",
                "Partial",
                "Complete"
              ]
            },
            {
              "field": "survey_person",
              "label": "Surveyed by",
              "type": "text",
              "length": 100
            },
            {
              "field": "survey_title",
              "label": "Title",
              "type": "text",
              "length": 100
            },
            {
              "field": "survey_agency",
              "label": "Agency / organization",
              "type": "text",
              "length": 150
            },
            {
              "field": "survey_date",
              "label": "Survey date",
              "type": "date"
            },
            {
              "field": "initial_survey",
              "label": "Initial survey",
              "type": "yesno",
              "length": 3
            },
            {
              "field": "resurvey",
              "label": "Re-survey",
              "type": "yesno",
              "length": 3
            },
            {
              "field": "notes",
              "label": "Notes / additional information",
              "type": "textarea",
              "length": 2000
            }
          ]
        }
      ],
    },
  ],
};
