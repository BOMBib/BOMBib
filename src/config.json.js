/* exported config */
var config = {
    "librarypath": "../library/projects/library.json",
    "categorys": [
        "Resistor",
        "Capacitor",
        "Potentiometer",
        "Diode",
        "IC",
        "Transistor",
    ],
    "category_spec_hints": {
        "Potentiometer": [
            "Mono", "Stereo",
            "Linear", "Logarithmic",
        ],
        "Resistor": [
            "Matched",
        ],
        "Capacitor": [
            "Ceramic",
            "Foil",
            "Electrolytic",
        ],
    },
    "projecttags": [
        "Osc",
        "VCO",
        "Filt",
        "VCF",
        "VCA",
        "Env",
        "ADSR",
        "AD",
        "Noise",
        "Seq",
        "Effect",
        "Util",
        "PSU",
        "-12,0,12V",
        "0,5V",
    ],
};
