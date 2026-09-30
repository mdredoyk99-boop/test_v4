"""Central configuration: dropdown options -> template tokens, Excel column aliases, naming."""

# ---------------------------------------------------------------- dropdowns
# label shown in the UI  ->  token key used in the template as {{CHK_<KEY>}}
PROTOCOLS = {
    "Knit Fabric 1": "PROTO_KNIT_FABRIC_1",
    "Knit Fabric 1A": "PROTO_KNIT_FABRIC_1A",
    "Knit Fabric (Additional Color) 1": "PROTO_KNIT_FABRIC_ADD_1",
    "Knit Garment 1": "PROTO_KNIT_GARMENT_1",
    "Knit Garment 1A": "PROTO_KNIT_GARMENT_1A",
    "Knit Garment (Additional Color) 1": "PROTO_KNIT_GARMENT_ADD_1",
    "Woven Garments (Solid, AOP, Yarn Dyed) 2": "PROTO_WOVEN_2",
    "Woven Garments (Solid, AOP, Yarn Dyed) 2A": "PROTO_WOVEN_2A",
    "Woven Garments (Solid, AOP, Yarn Dyed) (Additional Pkg. 2)": "PROTO_WOVEN_ADD_2",
    "Woven Garments (Garments Overdye, Denim wash) 3": "PROTO_WOVEN_3",
    "Woven Garments (Garments Overdye, Denim wash) 3A": "PROTO_WOVEN_3A",
    "Woven Garments (Garments Overdye, Denim wash) (Additional Pkg. 3)": "PROTO_WOVEN_ADD_3",
    "Outerwear Jacket 4": "PROTO_OUTER_4",
    "Outerwear Jacket 4A": "PROTO_OUTER_4A",
    "Outerwear Jacket (Additional Pkg. 4)": "PROTO_OUTER_ADD_4",
    "Sweater (Full Package) 1": "PROTO_SWEATER_1",
    "Sweater (Additional Color Package) 2": "PROTO_SWEATER_2",
    "Socks/Beanies (Full Package) 1": "PROTO_SOCKS_1",
    "Socks/Beanies (Additional Colorway) 2": "PROTO_SOCKS_2",
}

STAGES = {
    "Development": "STAGE_DEVELOPMENT",
    "TPA/SMS": "STAGE_TPA_SMS",
    "Bulk Production": "STAGE_BULK",
    "Sealed Sample": "STAGE_SEALED",
}

GENDERS = {"Men": "GENDER_MEN", "Women": "GENDER_WOMEN",
           "Children": "GENDER_CHILDREN", "Infant": "GENDER_INFANT"}
WASH = {"Yes": "WASH_YES", "No": "WASH_NO"}

# Used by prepare_template.py: template label (letters/digits only, lower-case) -> token key.
# Each entry is a regex that must match the whole normalised label.
PROTOCOL_LABEL_PATTERNS = [
    (r"knitfabric1", "PROTO_KNIT_FABRIC_1"),
    (r"knitfabric1a", "PROTO_KNIT_FABRIC_1A"),
    (r"knitfabricadditionalcolor1", "PROTO_KNIT_FABRIC_ADD_1"),
    (r"knitgarment1", "PROTO_KNIT_GARMENT_1"),
    (r"knitgarment1a", "PROTO_KNIT_GARMENT_1A"),
    (r"knitgarmentadditionalcolor1", "PROTO_KNIT_GARMENT_ADD_1"),
    (r"wovengarmentssoild.*pants2", "PROTO_WOVEN_2"),
    (r"wovengarmentssoild.*pants2a", "PROTO_WOVEN_2A"),
    (r"wovengarmentssoild.*additionalpkg2", "PROTO_WOVEN_ADD_2"),
    (r"wovengarmentsgarmentsoverdye.*jac3", "PROTO_WOVEN_3"),
    (r"wovengarmentsgarmentsoverdye.*jac3a", "PROTO_WOVEN_3A"),
    (r"wovengarmentsgarmentsoverdye.*additionalpkg3", "PROTO_WOVEN_ADD_3"),
    (r"outerwearjacket4", "PROTO_OUTER_4"),
    (r"outerwearjacket4a", "PROTO_OUTER_4A"),
    (r"outerwearjacketadditionalpkg4", "PROTO_OUTER_ADD_4"),
    (r"sweaterfullpackage1", "PROTO_SWEATER_1"),
    (r"sweateradditionalcolorpackage2", "PROTO_SWEATER_2"),
    (r"socksbeaniesfullpackage1", "PROTO_SOCKS_1"),
    (r"socksbeaniesadditionalcolorway2", "PROTO_SOCKS_2"),
]
STAGE_LABEL_PATTERNS = [
    (r"development", "STAGE_DEVELOPMENT"),
    (r"tpasms", "STAGE_TPA_SMS"),
    (r"bulkproduction", "STAGE_BULK"),
    (r"sealedsample", "STAGE_SEALED"),
]
GENDER_LABEL_PATTERNS = [(r"men", "GENDER_MEN"), (r"women", "GENDER_WOMEN"),
                         (r"children", "GENDER_CHILDREN"), (r"infant", "GENDER_INFANT")]
WASH_LABEL_PATTERNS = [(r"yes", "WASH_YES"), (r"no", "WASH_NO")]

# Text placeholders in the tagged template  ->  label of the row whose value cell they replace
TEXT_FIELDS = {
    "SEASON": "*Season:",
    "COLOR": "Color",
    "PROGRAM": "*Program#",
    "FIBRE": "*Fibre content",
    "STYLE": "Style #:",
    "CONSTRUCTION": "Construction",
    "PO": "*PO#",
    "GSM": "*Fabric/garment weight",   # cell text becomes "{{GSM}} GSM"
}
CARE_IMAGE_TOKEN = "CARE_IMAGE"

# ---------------------------------------------------------------- symbols
CHECKED = "\u2611"    # ☑
UNCHECKED = "\u2610"  # ☐
SYMBOL_FONT = "Segoe UI Symbol"

# ---------------------------------------------------------------- excel
COLUMN_ALIASES = {
    "program": ["PROG", "PROGRAM", "PROGRAMME", "PROGRAM#"],
    "style": ["STYLE", "STYLE NAME", "STYLE #"],
    "color": ["COLOR", "COLOUR", "COLORWAY"],
    "fabrication": ["FABRICATION", "FABRIC", "FABRIC DETAILS"],
    "gender": ["GENDER", "SEX"],
    "item_no": ["ITEM NO", "ITEM NUMBER", "ITEM #"],
    "wash": ["WASH TYPE", "WASH"],
    "po_candidates": ["TPA PO", "SMS PO", "HK PO", "PO", "PO#", "PO NO"],
}

# ---------------------------------------------------------------- output
# Placeholders: {style} {program} {color}. Underscores are always converted to spaces.
FILENAME_PATTERN = "{style} TRF {program}"
