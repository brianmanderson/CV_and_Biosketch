# -*- coding: utf-8 -*-
"""Hand-normalized author lists for every CV Publications entry dated 10/2022 or later.
Brian's own name is omitted from every list. Keys are 'Surname Initials' as printed in the CV
(after obvious typo fixes noted inline)."""
ENTRIES = [
 # (id, kind, short title, venue, date, [authors excluding Brian])
 ("E119","Book chapter","Artificial Intelligence in Adaptive Radiation Therapy, Ch. 9: Imaging Registration and Auto-Segmentation","Book (organized by Yi Wang, MGH)","undated",
  ["Brock KK"]),
 ("E122","Paper","Deep Learning-Based Vessel Auto-Segmentation for Standardized Lymph Node Delineation in Prostate Cancer Radiotherapy","Practical Radiation Oncology","09/2026",
  ["Gao Y","Wells D","Bhattacharya S","Repka M","Lavrova L","Das S","Rajasekar A","Byrd E","McGurk R","Dooley J","Wright J","Marks L","Mazur L"]),  # CV prints "Write J." (typo for Wright J.)
 ("E123","Paper","Using a Deep Learning System to Delineate the Elective Pelvis Nodal Volumes in Patients with Prostate Cancer","Advances in Radiation Oncology","07/2026",
  ["Repka M","Gao Y","Byrd E","McGurk R","Das S","Arcot R","Bhattacharya S","Adapa K","Wright J","Marks L","Mazur L"]),
 ("E124","Paper","Large language models for scalable standardization in radiation oncology","Physics and Imaging in Radiation Oncology","04/2026",
  ["Floyd W","Wahid KA","Knausl B","Moreno AC"]),
 ("E125","Preprint","Quality versus quantity of training datasets for artificial intelligence-based whole liver segmentation","medRxiv","02/2026",
  ["Castelo A","O'Connor C","Gupta A","Woodland M","Altaie M","Koay E","Odisio B","Tang T","Brock K"]),
 ("E126","Paper","Towards better understanding of factors contributing to medical physicist well-being in academic medical centers","Journal of Applied Clinical Medical Physics","05/2025",
  ["Kwong E","Liu CC","Adapa K","Vizer L","McHugh D","Pawlicki T","Miften M","Sawant A","Charguia N","Das S","Marks LB","Wright JL","Mazur L"]),
 ("E127","Paper","Software-based versus visual assessment of the minimal ablative margin (COVER-ALL): a randomised controlled phase 2 trial","The Lancet Gastroenterology & Hepatology","05/2025",
  ["Odisio BC","Albuquerque J","Lin YM","O'Connor CS","Rigaud B","Briones-Dimayuga M","Jones AK","Fellman BM","Huang SY","Kuban J","Metwalli ZA","Sheth R","Habibollahi P","Patel M","Shan KY","Cox VL","Kang HSC","Morris VK","Kopetz S","Javle MM","Kaseb A","Tzeng CW","Cao HT","Newhook T","Chun YS","Vauthey JN","Gupta S","Paolucci I","Brock KK"]),
 ("E128","Paper","DICOM Attribute Manipulation Tool","Journal of Applied Clinical Medical Physics","04/2025",
  ["Bojechko C"]),
 ("E129","Paper","Rapid in vivo EPID image prediction using analytically calculated attenuation and AI predicted scatter","Medical Physics","02/2025",
  ["Moore L","Bojechko C"]),
 ("E130","Paper","Open RT Structures: A Solution for TG-263 Accessibility","International Journal of Radiation Oncology, Biology, Physics","06/2024",
  ["Padilla L","Ryckman J","Covington E","Hong DS","Woods K","Katz MS","Estes C","Moore K","Bojechko C"]),
 ("E131","Paper","Perceptions on roadblocks to implementation of standardized nomenclature in radiation oncology: survey from TG-263U1","Journal of Applied Clinical Medical Physics","06/2024",
  ["Covington E","Suresh K","Barker M","Dess K","Price J","Moncio A","Vaccarelli M","Santanam L","Xiao Y","Mayo C"]),
 ("E132","Paper","Identifying the optimal deep learning architecture and parameters for automatic beam aperture definition in 3D radiotherapy","Journal of Applied Clinical Medical Physics","12/2023",
  ["Gay S","Kisling K","Zhang L","Rhee DJ","Nguyen C","Netherton T","Yang J","Brock K","Jhingran A","Simonds H","Klopp A","Beadle BM","Court L","Cardenas C"]),
 ("E133","Paper","Comparison of Machine-Learning and Deep-Learning Methods for the Prediction of Osteoradionecrosis","Advances in Radiation Oncology","07/2023",
  ["Reber B","Van Dijk L","Mohamed ASR","Fuller C","Lai S","Brock K"]),
 ("E134","Paper","Ablative Margins of Colorectal Liver Metastases Using Deformable CT Image Registration and Autosegmentation","Radiology","04/2023",
  ["Lin YM","Paolucci I","O'Connor CS","Rigaud B","Fellman BM","Jones KA","Brock KK","Odisio BC"]),
 ("E135","Paper","Deep Learning Models for Automated Assessment of Breast Density Using Multiple Mammographic Image Types","Cancers","10/2022",
  ["Rigaud B","Weaver OO","Dennison JB","Awais M","Chiang TYD","Yang WT","Hanash SM","Brock KK"]),
 ("E136","Paper","Leveraging deep learning-based segmentation and contours-driven deformable registration for dose accumulation in abdominal structures","Frontiers in Oncology","10/2022",
  ["McCulloch MM","Cazoulat G","Svensson S","Gryshkevych S","Rigaud B","Kirimli E","De B","Mathew RT","Zaid M","Elganainy D","Peterson CB","Balter P","Koay EJ","Brock KK"]),
 ("E165","Oral presentation","Enabling Adoption of TG-263 Standardization of Nomenclature: A Tool to Reduce the Headache","AAMD Annual Conference","06/2023",
  ["Moore K","Padilla L","Bojechko C"]),
 ("E178","Poster","Predicting Elective Nodal Volumes with Deep Learning: A Tool to Facilitate Peer Review","UNC Chapel Hill Research Retreat","05/2025",
  ["Foster M","Repka M","Karunaker A","Marks LB","Das S","Mazur L"]),
 ("E179","Poster","Simple Method for Functional Sparing of Parotid Glands in Head and Neck IMRT","ASTRO","09/2024",
  ["Fried D","Hawkins M","Das S","Repka M","Shen C","Chen X"]),   # CV prints "Cen X" (typo?)
 ("E180","Poster","Unlocking Insights: A Python-Based RayStation Data Mining Tool","AAPM","07/2024",
  ["Baughan N","Marks L","Chen S","Yanagihara T","Das S"]),
 ("E181","Poster","Predicting Dysphagia in Head and Neck Radiation Therapy with Deep Learning","AAPM","07/2024",
  ["Das S","Fried D","Dong F","Morse R","Shen C","Repka M","Chen S"]),
 ("E187","Abstract","Root Cause Analysis of Unscheduled Replans Across 19,249 Radiation Therapy Plans","ASTRO Annual Conference","09/2026",
  ["Abdou M","Stark J","Kim S","Zegers B"]),
 ("E188","Abstract","Prospective Assessment of the Impact of AI on Peer Review of Elective Pelvic Nodal Segmentations","ASTRO Annual Conference","09/2026",
  ["Repka M","Marks L","Byrd E","Das S","Gao Y","Adapa K","McGurk R","Mazur L"]),
 ("E189","Abstract","Deep Learning-Based Pelvic Vessel Auto-Segmentation for Standardized Lymph Node Delineation in Prostate Cancer Radiotherapy","ASTRO Annual Conference","09/2026",
  ["Gao Y","Wells D","Bhattacharya S","Repka MC","Lavrova L","Das SD","Rajasekar A","Byrd E","McGurk R","Dooley J","Adapa K","Wright JL","Marks LB","Mazur LM"]),
 ("E190","Abstract","Predicting Dysphagia and Identifying Anatomical Regions Responsible for Dysphagia Following IMRT for Oropharyngeal Carcinoma","ASTRO Annual Conference","09/2026",
  ["Das S","Dong F","Morse R","Chera B","Shen C","Repka M","Yarbrough W","Hackman T","Blumberg J","Lumley C","Weiss J","Patel S","Sheth S","Chen X"]),
 ("E191","Abstract","Cumulative EQD2 Esophageal Dose and Severe Toxicity After Thoracic Reirradiation","ASTRO Annual Conference","09/2026",
  ["Sousa C","Doss V","Hales E","Hagan E","Gebre T","Obaideen A","Negassa M","Liu D","Kooner","Hu C","Li H","Viswanathan A","Jia X","Hales R","Voong KR","McNutt T","Ger R"]),
 ("E192","Abstract","Head and Neck Gross Tumor Deep-Learning Autocontouring Model Using BigROC Collaborative Data","AAPM Annual Conference","07/2026",
  ["Abdulkadir Y","Fields E","Hink JB","Hsu D","Kisling K","Krupein A","Lamb J","Matney J","Maxin P","Moore L","Neylon J","Ritter T","Savjani R","Valdes G","Vasconcelos N","Vinogradskiy Y","Witztum A"]),
 ("E193","Abstract","Rapid Semi-Automated Planning for VMAT Craniospinal Irradiation Using Raystation","AAPM Annual Conference","07/2026",
  ["Friend D","Dance M","Dooley J","Casey D","Shen C","Das S"]),
 ("E194","Abstract","Automated Detection of Hip Replacements and Fiducial Markers in Pelvic CT Scans","IEEE Conference on AI","06/2026",
  ["Bhattacharya S","Karunaker A","Wells D","Gao Y","Marks L","Das S","Rajasekar A","Mazur L"]),
 ("E195","Abstract","Navigating Variability in Prostate RT Planning: Real-Time Insights for Human-Centered CDS Design","AMIA Annual Symposium","02/2026",
  ["Foster M","Byrd E","Kwong E","Karunaker A","Repka M","McGurk R","Das S","Marks L","Mazur L"]),
]

# Variant spellings that are the same person (canonical key -> variants)
MERGE = {
 "Brock KK": ["Brock K"],
 "Marks LB": ["Marks L"],
 "Mazur LM": ["Mazur L"],
 "Das SD": ["Das S"],
 "Repka MC": ["Repka M"],
 "Wright JL": ["Wright J"],
 "Jones AK": ["Jones KA"],
 "O'Connor CS": ["O'Connor C"],
 "Koay EJ": ["Koay E"],
 "Odisio BC": ["Odisio B"],
}
VARIANT_TO_CANON = {v: k for k, vs in MERGE.items() for v in vs}


def canon(name):
    return VARIANT_TO_CANON.get(name, name)


def roster():
    people = {}
    for eid, kind, title, venue, date, authors in ENTRIES:
        for a in authors:
            c = canon(a)
            people.setdefault(c, {"variants": set(), "entries": []})
            people[c]["variants"].add(a)
            people[c]["entries"].append((eid, kind, title, venue, date))
    return people


if __name__ == "__main__":
    ppl = roster()
    print("unique people:", len(ppl))
    for k in sorted(ppl):
        ents = ppl[k]["entries"]
        print(f"{k:22s} n={len(ents):2d}  " + "; ".join(f"{e[0]}({e[4]})" for e in ents))
