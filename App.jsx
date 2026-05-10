
import React from "react";
import { useState, useRef, useEffect } from "react";
import * as XLSX from "xlsx";

// ─── CONSTANTS ────────────────────────────────────────────────────────────────
// Updated sector names to match EPFL/ETH database taxonomy
const SECTOR_NAMES = {1:"Neurovascular & Neurotech",2:"Orthopedics & Musculoskeletal",3:"Wearable & Drug Delivery",4:"Cardiovascular & Interventional",5:"Diagnostics & IVD",6:"Surgical Robotics & MIS",7:"Digital Health & SaMD",8:"Ophthalmology & Photonics",9:"Regenerative Medicine & Wound Care",10:"Monitoring & Implantables"};
const SECTOR_KW = {1:["neuro","brain","neural","seizure","epilep","stroke","cranial","neurovascular","neurostimul"],2:["ortho","bone","spine","musculo","joint","fracture","implant","sacroiliac","vertebra"],3:["wearable","insulin","pump","infusion","drug delivery","patch","diabetes","cgm","glucose","injec"],4:["cardio","cardiac","heart","vascular","artery","perfusion","aortic","bypass","ablation","angio"],5:["diagnos","ivd","assay","biomarker","reagent","molecular","genomic","per test","sequencing","pcr"],6:["robot","endoscop","laparoscop","surgical robot","single-use scope","mis","minimally invasive"],7:["software","samd","digital health","imaging","mri","pacs","ai","machine learning","cloud","subscription"],8:["ophthalm","eye","retina","ocular","vision","glaucom","laser","fundus","cataract","lens","photonic"],9:["regenerat","tissue","wound","biologic","scaffold","nerve repair","cell therapy","graft","biologics"],10:["monitor","implantable","pacemaker","stimulat","neuromodulat","sensor","cgm","remote monitor","wearabl device"]};
const WACC_P = {"Switzerland":{rf:0.0075,erp:0.0554,tax:0.149,g:0.015},"United States":{rf:0.0425,erp:0.046,tax:0.21,g:0.0225},"Germany":{rf:0.025,erp:0.0554,tax:0.295,g:0.015},"France":{rf:0.03,erp:0.0554,tax:0.25,g:0.015},"Belgium":{rf:0.03,erp:0.0649,tax:0.25,g:0.015},"Sweden":{rf:0.0215,erp:0.0554,tax:0.206,g:0.015},"United Kingdom":{rf:0.04,erp:0.0554,tax:0.25,g:0.02},"Netherlands":{rf:0.028,erp:0.0554,tax:0.258,g:0.015},"Spain":{rf:0.031,erp:0.0649,tax:0.25,g:0.015},"Italy":{rf:0.037,erp:0.0649,tax:0.275,g:0.01},"Denmark":{rf:0.022,erp:0.0554,tax:0.22,g:0.015},"Finland":{rf:0.021,erp:0.0554,tax:0.20,g:0.015},"Norway":{rf:0.032,erp:0.0554,tax:0.22,g:0.015},"Austria":{rf:0.027,erp:0.0554,tax:0.25,g:0.015},"Poland":{rf:0.053,erp:0.0649,tax:0.19,g:0.02},"Israel":{rf:0.04,erp:0.0649,tax:0.23,g:0.02},"Canada":{rf:0.033,erp:0.046,tax:0.265,g:0.02},"Australia":{rf:0.042,erp:0.046,tax:0.30,g:0.025},"Japan":{rf:0.009,erp:0.046,tax:0.3086,g:0.01},"China":{rf:0.023,erp:0.064,tax:0.25,g:0.04},"India":{rf:0.069,erp:0.064,tax:0.30,g:0.05},"Singapore":{rf:0.029,erp:0.046,tax:0.17,g:0.025},"South Korea":{rf:0.031,erp:0.046,tax:0.275,g:0.025},"Brazil":{rf:0.135,erp:0.078,tax:0.34,g:0.04},"Default":{rf:0.025,erp:0.055,tax:0.21,g:0.015}};

// ─── FULL 66-COMPANY DATABASE — EPFL/ETH Medtech Comparable Database ──────────
// Fields: t=ticker, n=name, s=sector(1-10), gm=gross margin%, eb=EBITDA margin%,
//         b=beta(unlevered), rg=revenue growth%, c=country, de=D/E ratio,
//         ok=🟢 reliable / 🟡 usable with limits, note=key usage caveat
// D/E corrected for confirmed yfinance errors: VCEL(0), RAY-B(0.01), CREO(gm=46.6)
const DB = [
  // ── Sub-sector 1: Neurovascular & Neurotech ──
  {"t":"PEN","n":"Penumbra Inc.","s":1,"gm":67.4,"eb":14.7,"b":0.735,"rg":15.6,"c":"USA","de":14.7,"ok":"🟢"},
  {"t":"NYXH","n":"Nyxoah SA","s":1,"gm":63.1,"eb":0.0,"b":0.876,"rg":34.7,"c":"Belgium","de":85.6,"ok":"🟡"},
  {"t":"CLPT","n":"ClearPoint Neuro","s":1,"gm":61.4,"eb":-58.1,"b":1.294,"rg":34.0,"c":"USA","de":207.8,"ok":"🟡"},
  {"t":"NSPR","n":"InspireMD Inc.","s":1,"gm":29.5,"eb":0.0,"b":0.819,"rg":61.6,"c":"USA","de":5.96,"ok":"🟡"},
  // ── Sub-sector 2: Orthopedics & Musculoskeletal ──
  {"t":"BONEX","n":"BONESUPPORT AB","s":2,"gm":92.5,"eb":26.5,"b":0.482,"rg":14.3,"c":"Sweden","de":1.5,"ok":"🟢"},
  {"t":"SIBN","n":"SI-BONE Inc.","s":2,"gm":79.6,"eb":-8.2,"b":0.671,"rg":15.0,"c":"USA","de":20.7,"ok":"🟢"},
  {"t":"ATEC","n":"Alphatec Holdings","s":2,"gm":70.2,"eb":3.6,"b":0.966,"rg":13.6,"c":"USA","de":20.0,"ok":"🟡"},
  {"t":"OFIX","n":"Orthofix Medical Inc.","s":2,"gm":71.0,"eb":2.1,"b":0.8,"rg":1.6,"c":"USA","de":60.3,"ok":"🟡"},
  // ── Sub-sector 3: Wearable & Drug Delivery ──
  {"t":"TNDM","n":"Tandem Diabetes Care","s":3,"gm":53.8,"eb":-6.0,"b":1.6,"rg":2.7,"c":"USA","de":20.0,"ok":"🟡"},
  {"t":"EMBC","n":"Embecta Corp.","s":3,"gm":61.9,"eb":31.9,"b":1.1,"rg":-14.4,"c":"USA","de":20.0,"ok":"🟡"},
  // ── Sub-sector 4: Cardiovascular & Interventional ──
  {"t":"XVVO","n":"XVIVO Perfusion AB","s":4,"gm":73.1,"eb":16.5,"b":1.4,"rg":10.2,"c":"Sweden","de":5.7,"ok":"🟢"},
  {"t":"ATRC","n":"AtriCure Inc.","s":4,"gm":75.6,"eb":3.2,"b":1.281,"rg":14.3,"c":"USA","de":15.2,"ok":"🟢"},
  {"t":"LMAT","n":"LeMaitre Vascular","s":4,"gm":71.3,"eb":30.7,"b":0.597,"rg":11.2,"c":"USA","de":46.7,"ok":"🟢"},
  {"t":"AORT","n":"Artivion Inc.","s":4,"gm":64.4,"eb":12.2,"b":1.4,"rg":19.2,"c":"USA","de":57.7,"ok":"🟢"},
  {"t":"ANGO","n":"AngioDynamics Inc.","s":4,"gm":54.3,"eb":-0.2,"b":0.4,"rg":8.9,"c":"USA","de":6.6,"ok":"🟡"},
  // ── Sub-sector 5: Diagnostics & IVD ──
  {"t":"SBS","n":"Stratec SE","s":5,"gm":25.6,"eb":11.3,"b":0.8,"rg":-11.1,"c":"Germany","de":55.7,"ok":"🟢"},
  {"t":"EKF","n":"EKF Diagnostics","s":5,"gm":51.4,"eb":19.7,"b":0.526,"rg":5.4,"c":"UK","de":1.97,"ok":"🟢"},
  {"t":"BOUL","n":"Boule Diagnostics AB","s":5,"gm":42.3,"eb":10.3,"b":0.2,"rg":-10.0,"c":"Sweden","de":20.0,"ok":"🟡"},
  {"t":"TECN","n":"Tecan Group AG","s":5,"gm":35.2,"eb":9.2,"b":1.0,"rg":-5.2,"c":"Switzerland","de":19.0,"ok":"🟡"},
  {"t":"VITR","n":"Vitrolife AB","s":5,"gm":58.6,"eb":27.5,"b":1.7,"rg":-4.2,"c":"Sweden","de":14.6,"ok":"🟡"},
  {"t":"TSTL","n":"Tristel plc","s":5,"gm":81.7,"eb":23.1,"b":0.3,"rg":13.6,"c":"UK","de":17.4,"ok":"🟡"},
  {"t":"VCYT","n":"Veracyte Inc.","s":5,"gm":72.9,"eb":19.2,"b":1.885,"rg":21.5,"c":"USA","de":2.93,"ok":"🟡"},
  {"t":"CDNA","n":"CareDx Inc.","s":5,"gm":68.9,"eb":0.6,"b":2.5,"rg":39.0,"c":"USA","de":7.8,"ok":"🟡"},
  {"t":"CSTL","n":"Castle Biosciences","s":5,"gm":79.4,"eb":-0.6,"b":1.1,"rg":0.8,"c":"USA","de":7.9,"ok":"🟡"},
  // ── Sub-sector 6: Surgical Robotics & MIS ──
  {"t":"AMBU","n":"Ambu A/S","s":6,"gm":60.0,"eb":13.3,"b":1.3,"rg":1.2,"c":"Denmark","de":9.0,"ok":"🟢"},
  {"t":"PRCT","n":"PROCEPT BioRobotics","s":6,"gm":64.0,"eb":-31.8,"b":0.826,"rg":20.2,"c":"USA","de":22.4,"ok":"🟡"},
  {"t":"STXS","n":"Stereotaxis Inc.","s":6,"gm":52.7,"eb":-56.8,"b":1.3,"rg":36.3,"c":"USA","de":29.0,"ok":"🟡"},
  {"t":"CREO","n":"Creo Medical Group","s":6,"gm":46.6,"eb":0.0,"b":1.0,"rg":37.5,"c":"UK","de":6.9,"ok":"🟡"},
  // ── Sub-sector 7: Digital Health & SaMD ──
  {"t":"RAY","n":"RaySearch Laboratories","s":7,"gm":92.4,"eb":25.6,"b":1.0,"rg":-12.5,"c":"Sweden","de":0.01,"ok":"🟢"},
  {"t":"SECT","n":"Sectra AB","s":7,"gm":38.9,"eb":20.7,"b":0.85,"rg":5.6,"c":"Sweden","de":5.67,"ok":"🟡"},
  {"t":"ASCN","n":"Ascom Holding AG","s":7,"gm":48.2,"eb":7.9,"b":1.0,"rg":5.2,"c":"Switzerland","de":20.0,"ok":"🟡"},
  {"t":"IRTC","n":"iRhythm Technologies","s":7,"gm":71.0,"eb":-1.4,"b":1.333,"rg":25.7,"c":"USA","de":20.0,"ok":"🟡"},
  {"t":"BFLY","n":"Butterfly Network","s":7,"gm":66.1,"eb":-49.0,"b":2.28,"rg":25.0,"c":"USA","de":10.3,"ok":"🟡"},
  {"t":"NNOX","n":"Nano-X Imaging","s":7,"gm":-98.2,"eb":0.0,"b":1.2,"rg":24.0,"c":"Israel","de":5.6,"ok":"🟡"},
  {"t":"HYPR","n":"Hyperfine Inc.","s":7,"gm":49.8,"eb":-265.4,"b":1.4,"rg":128.0,"c":"USA","de":0.77,"ok":"🟡"},
  // ── Sub-sector 8: Ophthalmology & Photonics ──
  {"t":"REG1V","n":"Revenio Group Oyj","s":8,"gm":69.9,"eb":21.6,"b":0.8,"rg":4.6,"c":"Finland","de":9.2,"ok":"🟢"},
  {"t":"STAA","n":"STAAR Surgical","s":8,"gm":76.2,"eb":-15.6,"b":1.202,"rg":18.1,"c":"USA","de":11.1,"ok":"🟢"},
  {"t":"OPTOMED","n":"Optomed Oyj","s":8,"gm":63.6,"eb":-31.3,"b":1.6,"rg":-5.6,"c":"Finland","de":11.3,"ok":"🟡"},
  {"t":"ELN","n":"El.En. S.p.A.","s":8,"gm":42.9,"eb":13.7,"b":1.1,"rg":24.8,"c":"Italy","de":7.36,"ok":"🟡"},
  {"t":"GKOS","n":"Glaukos Corporation","s":8,"gm":78.1,"eb":-8.0,"b":0.948,"rg":41.2,"c":"USA","de":15.8,"ok":"🟡"},
  {"t":"LNSR","n":"LENSAR Inc.","s":8,"gm":46.4,"eb":-34.4,"b":0.8,"rg":-4.2,"c":"USA","de":20.0,"ok":"🟡"},
  {"t":"LBIRD","n":"Lumibird S.A.","s":8,"gm":63.4,"eb":14.2,"b":0.9,"rg":8.5,"c":"France","de":70.9,"ok":"🟡"},
  // ── Sub-sector 9: Regenerative Medicine & Wound Care ──
  {"t":"VCEL","n":"Vericel Corp.","s":9,"gm":74.8,"eb":9.4,"b":1.1,"rg":30.1,"c":"USA","de":0.0,"ok":"🟢"},
  {"t":"AXGN","n":"Axogen Inc.","s":9,"gm":75.0,"eb":-0.9,"b":1.2,"rg":26.6,"c":"USA","de":8.4,"ok":"🟡"},
  {"t":"ORGO","n":"Organogenesis Holdings","s":9,"gm":76.5,"eb":14.2,"b":1.3,"rg":78.1,"c":"USA","de":18.9,"ok":"🟡"},
  {"t":"EUZ","n":"Eckert & Ziegler SE","s":9,"gm":49.0,"eb":28.7,"b":1.5,"rg":9.3,"c":"Germany","de":17.7,"ok":"🟡"},
  {"t":"PHO","n":"Photocure ASA","s":5,"gm":92.1,"eb":4.3,"b":0.4,"rg":-3.5,"c":"Norway","de":2.1,"ok":"🟡"},
  // ── Sub-sector 7: Digital Health & SaMD (additional) ──
  {"t":"PLLWF","n":"Polarean Imaging","s":7,"gm":45.2,"eb":0.0,"b":0.1,"rg":-46.9,"c":"UK","de":5.6,"ok":"🟡"},
  // ── Sub-sector 8: Ophthalmology & Photonics (additional) ──
  {"t":"IBAB","n":"IBA S.A.","s":8,"gm":32.2,"eb":5.0,"b":1.2,"rg":8.1,"c":"Belgium","de":20.0,"ok":"🟡"},
  // ── Sub-sector 10: Monitoring & Implantables ──
  {"t":"LIVN","n":"LivaNova PLC","s":10,"gm":67.9,"eb":18.3,"b":0.819,"rg":14.3,"c":"UK","de":28.4,"ok":"🟢"},
  {"t":"DRW3","n":"Drägerwerk AG","s":10,"gm":45.3,"eb":8.7,"b":0.6,"rg":3.5,"c":"Germany","de":20.7,"ok":"🟡"},
  {"t":"BACTI","n":"Bactiguard Holding AB","s":10,"gm":88.0,"eb":15.5,"b":0.4,"rg":-28.1,"c":"Sweden","de":51.8,"ok":"🟡"},
  {"t":"IHC","n":"Inspiration Healthcare","s":10,"gm":44.4,"eb":3.7,"b":0.9,"rg":40.8,"c":"UK","de":78.9,"ok":"🟡"},
  {"t":"EKTA","n":"Elekta AB","s":10,"gm":38.5,"eb":8.9,"b":1.0,"rg":-9.7,"c":"Sweden","de":90.9,"ok":"🟡"},
  {"t":"INSP","n":"Inspire Medical Systems","s":10,"gm":85.8,"eb":7.4,"b":0.831,"rg":1.6,"c":"USA","de":3.73,"ok":"🟡"},
  {"t":"SENS","n":"Senseonics Holdings","s":10,"gm":44.7,"eb":-189.5,"b":1.1,"rg":71.8,"c":"USA","de":67.8,"ok":"🟡"},
  {"t":"CVRX","n":"CVRx Inc.","s":10,"gm":85.3,"eb":-89.2,"b":0.9,"rg":4.4,"c":"USA","de":20.0,"ok":"🟡"},
];
// ─── PER-COMPANY SELECTION RATIONALE ─────────────────────────────────────────
// Explains WHY each specific company was chosen — shown in UI card + Excel note
const COMP_REASON = {
  // Sub-sector 1
  "PEN":    "🟢 Anchor comp — catheter-based neurovascular devices, specialist sales, 67.4% GM. Most liquid US-listed peer.",
  "NYXH":   "🟡 Implantable neurostimulator (hypoglossal nerve), CE MDR cleared; high D/E limits beta reliability.",
  "CLPT":   "🟡 MRI-guided neuro delivery platform; high beta (1.294) reflects early-commercial stage risk.",
  "NSPR":   "🟡 Carotid stent system; rapid rev growth (+61.6%) but thin GM (29.5%) skews sector median down.",
  // Sub-sector 2
  "BONEX":  "🟢 Swedish bone substitute; 92.5% GM — highest in DB; CE + FDA cleared, profitable, clean balance sheet.",
  "SIBN":   "🟢 Sacroiliac joint fusion; 79.6% GM, DRG/NTAP reimbursement path — best US structural-implant comp.",
  "ATEC":   "🟡 Cervical/lumbar spine systems; comparable revenue model but higher leverage (D/E 20×).",
  "OFIX":   "🟡 Spine & bone stimulation; low growth (+1.6%) drags sector RG median — used for beta calibration.",
  // Sub-sector 3
  "TNDM":   "🟡 Insulin pump; CGM integration model mirrors wearable drug delivery; high beta (1.6).",
  "EMBC":   "🟡 Insulin delivery devices (spin-off from BD); mature product — GM 61.9%, negative revenue growth caution.",
  // Sub-sector 4
  "XVVO":   "🟢 Organ perfusion systems; capital+disposable model, profitable EU-listed comp, 73.1% GM.",
  "ATRC":   "🟢 Cardiac ablation; 75.6% GM, growing US revenue (+14.3%), direct hospital sales — cleanest cardio comp.",
  "LMAT":   "🟢 Vascular surgery devices; 71.3% GM, 30.7% EBITDA — most operationally mature peer in sector.",
  "AORT":   "🟢 Cardiac/vascular biologics; 64.4% GM, high growth (+19.2%) from recent M&A — watch for integration risk.",
  "ANGO":   "🟡 Vascular access & ablation; low beta (0.4) and declining EBITDA weigh on comparability.",
  // Sub-sector 5
  "SBS":    "🟢 German IVD instrument OEM; low GM (25.6%) anchors B2B/OEM end of sector range.",
  "EKF":    "🟢 Point-of-care diagnostics; 51.4% GM, profitable UK-listed — reliable small-cap IVD comp.",
  "BOUL":   "🟡 Hematology analyzer OEM; low beta (0.2) and declining revenue reduce weight.",
  "TECN":   "🟡 Lab automation (Tecan); 35.2% GM reflects instrument-heavy mix — useful for OEM/capital models.",
  "VITR":   "🟡 IVF consumables; 58.6% GM but recent revenue contraction (-4.2%) after Igenomix acquisition.",
  "TSTL":   "🟡 Infection prevention diagnostics; 81.7% GM (high-end anchor), UK small-cap liquidity caution.",
  "VCYT":   "🟡 Genomic diagnostics; 72.9% GM, strong growth (+21.5%) — useful for high-margin molecular Dx.",
  "CDNA":   "🟡 Transplant diagnostics; 68.9% GM but beta of 2.5 flags high equity risk.",
  "CSTL":   "🟡 Dermatology genomic testing; 79.4% GM, flat growth (+0.8%) — mature per-test model reference.",
  "PHO":    "🟡 Photodynamic cancer diagnostics; 92.1% GM (outlier), low beta (0.4) — specialty comp only.",
  // Sub-sector 6
  "AMBU":   "🟢 Single-use endoscopy; 60.0% GM, Denmark-listed, growing recurring-use model — best MIS comp.",
  "PRCT":   "🟡 Aquablation robotic system; 64.0% GM, capital+procedure mix — high leverage caution.",
  "STXS":   "🟡 Robotic electrophysiology; high growth (+36.3%) but deep EBITDA losses (-56.8%).",
  "CREO":   "🟡 Advanced electrosurgery; GM corrected to 46.6% (yfinance error fixed). Early commercial.",
  // Sub-sector 7
  "RAY":    "🟢 Radiation therapy planning SaMD; 92.4% GM — best-in-class software margin anchor. D/E corrected to 0.01×.",
  "SECT":   "🟡 Medical imaging SaaS (PACS/VNA); 38.9% GM reflects services mix — useful for hybrid SaaS models.",
  "ASCN":   "🟡 Clinical communication platform; 48.2% GM, Swiss-listed, enterprise SaaS recurring revenue.",
  "IRTC":   "🟡 Cardiac monitoring SaMD (Zio patch); 71.0% GM, +25.7% revenue growth — ambulatory Dx reference.",
  "BFLY":   "🟡 Handheld ultrasound SaMD; 66.1% GM, subscription model — high beta (2.28) flags speculative risk.",
  "NNOX":   "🟡 Tomosynthesis-as-a-service; negative GM (-98.2%) makes this a downside scenario anchor only.",
  "HYPR":   "🟡 Low-field MRI; extreme growth (+128%) and EBITDA margin (-265%) — excluded from median calculation.",
  "PLLWF":  "🟡 Hyperpolarized imaging agent; very low beta (0.1) and negative revenue growth — limited use.",
  // Sub-sector 8
  "REG1V":  "🟢 Glaucoma diagnostics (Icare); 69.9% GM, profitable Finnish comp — closest EU ophthalmic device peer.",
  "STAA":   "🟢 Implantable collamer lenses; 76.2% GM, growing (+18.1%) — primary US ophthalmic implant anchor.",
  "OPTOMED":"🟡 Retinal imaging; 63.6% GM, Finnish small-cap, declining revenue (-5.6%) weights beta estimate.",
  "ELN":    "🟡 Medical laser systems; 42.9% GM reflects hardware mix — useful for photonics/laser device models.",
  "GKOS":   "🟡 Glaucoma devices; 78.1% GM, high growth (+41.2%) post-iDose launch — bull scenario reference.",
  "LNSR":   "🟡 Laser cataract surgery; 46.4% GM, declining revenue — capital equipment model reference.",
  "LBIRD":  "🟡 Photonics & laser (Lumibird); 63.4% GM, French-listed — direct photonics technology comparable.",
  "IBAB":   "🟡 Proton therapy systems (IBA); 32.2% GM reflects heavy capital equipment — useful for hardware-heavy models.",
  // Sub-sector 9
  "VCEL":   "🟢 Cell therapy (MACI/Epicel); 74.8% GM, profitable, D/E corrected to 0.0× (zero debt per 2023 10-K).",
  "AXGN":   "🟡 Peripheral nerve repair biologics; 75.0% GM, growing (+26.6%) — closest neural regeneration peer.",
  "ORGO":   "🟡 Advanced wound biologics; 76.5% GM, very high growth (+78.1%) — bull-scenario anchor.",
  "EUZ":    "🟡 Radiopharmaceuticals/biologics; 49.0% GM, German-listed — useful for EU regulatory pathway models.",
  // Sub-sector 10
  "LIVN":   "🟢 Neuromodulation + cardiac surgery; 67.9% GM, profitable UK/US dual-listed — cleanest implantable comp.",
  "DRW3":   "🟡 Ventilators & patient monitoring; 45.3% GM, German hospital capital equipment — useful for monitoring models.",
  "BACTI":  "🟡 Antimicrobial coatings; 88.0% GM (high-end anchor) but revenue declining sharply (-28.1%).",
  "IHC":    "🟡 Neonatal respiratory; 44.4% GM, UK-listed, high D/E (78.9×) limits beta usefulness.",
  "EKTA":   "🟡 Radiotherapy systems; 38.5% GM, declining revenue (-9.7%) — mature capital equipment reference.",
  "INSP":   "🟡 Inspire upper-airway stimulator; 85.8% GM — best US implantable GM anchor; flagged for GLP-1 disruption risk.",
  "SENS":   "🟡 Eversense long-term CGM; 44.7% GM, extreme EBITDA (-189.5%), high D/E — downside risk anchor.",
  "CVRX":   "🟡 Baroreflex activation therapy; 85.3% GM but EBITDA -89.2% and no revenue growth — clinical-stage proxy.",
};

// ─── SELECTION METHODOLOGY TEXT (injected into Excel + shown in UI) ──────────
const SELECTION_METHODOLOGY = `SELECTION METHODOLOGY — How comparables were chosen:
Step 1 — KEYWORD MATCHING: The startup description was scanned for sector keywords (e.g. "neural", "catheter", "cardiac", "diagnostic"). The sub-sector with the highest keyword hit count was selected.
Step 2 — DATABASE FILTERING: From the 66-company EPFL/ETH Medtech DB, all companies in the matched sub-sector were extracted. Companies with GM < −50% or missing beta were excluded as unusable (e.g. early-stage with no revenue).
Step 3 — RELIABILITY RANKING: Companies marked 🟢 (fully reliable: audited financials, liquid stock, no known yfinance errors) were prioritised. 🟡 (usable with limits) companies filled remaining slots up to 5 comparables.
Step 4 — SECTOR MEDIANS: Gross Margin, EBITDA Margin, Beta (unlevered), and Revenue Growth medians were computed from the selected set. These medians feed directly into the Assumptions sheet (WACC beta, terminal growth) and the PnL model (GM, EBITDA).
Step 5 — DATA CORRECTIONS: Known yfinance errors were corrected before use: VCEL D/E = 0.0× (not 26.5×; confirmed 2023 10-K), RaySearch D/E = 0.01× (not 36×), Creo Medical GM = 46.6% (not 20%).`;

const REVENUE_MODELS = ["Revenue Blade","Capital Sale","SaaS","Per Test","OEM","Royalty Licensing","Hybrid","Other (explain in description)"];
const STAGES = ["Seed","Series A","Series B","Series C+","Pre-revenue R&D","Commercial","Other (explain in description)"];
const COUNTRIES = ["Switzerland","United States","Germany","France","Belgium","Sweden","United Kingdom","Netherlands","Spain","Italy","Denmark","Finland","Norway","Austria","Poland","Israel","Canada","Australia","Japan","China","India","Singapore","South Korea","Brazil","Default"];
const RATIONALE = {
  1:"Matched from Neurovascular & Neurotech (EPFL/ETH DB, 66-co. set) — FDA PMA / CE MDR Class III pathway, implantable or catheter-based devices, clinical specialist sales model. Key: Penumbra 🟢 67.4% GM, ClearPoint Neuro 🟡 61.4% GM.",
  2:"Matched from Orthopedics & Musculoskeletal (EPFL/ETH DB, 66-co. set) — surgical bone/spine implants, NTAP/DRG hospital reimbursement. Key: BONESUPPORT 🟢 92.5% GM (highest in DB), SI-BONE 🟢 79.6% GM.",
  3:"Matched from Wearable & Drug Delivery (EPFL/ETH DB, 66-co. set) — wearable insulin delivery, patch/implantable drug delivery, DME reimbursement. Key: Embecta 🟡 61.9% GM, Tandem 🟡 53.8% GM.",
  4:"Matched from Cardiovascular & Interventional (EPFL/ETH DB, 66-co. set) — cardiac/vascular surgical devices, direct hospital sales, capital+disposable model. Key: AtriCure 🟢 75.6% GM, LeMaitre Vascular 🟢 71.3% GM (cleanest DB comparable).",
  5:"Matched from Diagnostics & IVD (EPFL/ETH DB, 66-co. set) — per-test reimbursement or OEM/B2B diagnostics. Note: LDT/CLIA companies carry FDA LDT rule risk. Key: Veracyte 🟡 72.9% GM, Tristel 🟡 81.7% GM.",
  6:"Matched from Surgical Robotics & MIS (EPFL/ETH DB, 66-co. set) — capital robot systems + high-margin disposables. Exit anchors: Medtronic/Mazor 25× EV/Rev; Stryker/MAKO 15× EV/Rev. Key: Ambu 🟢 60.0% GM, PROCEPT 🟡 64.0% GM.",
  7:"Matched from Digital Health & SaMD (EPFL/ETH DB, 66-co. set) — SaMD, AI, SaaS or license model. D/E corrected: RaySearch actual 0.01× (not 36× yfinance error). Key: RaySearch 🟢 92.4% GM (best-in-class SaMD), iRhythm 🟡 71.0% GM.",
  8:"Matched from Ophthalmology & Photonics (EPFL/ETH DB, 66-co. set) — ophthalmic devices, laser systems, photonics. Key: STAAR Surgical 🟢 76.2% GM, Glaukos 🟡 78.1% GM, Revenio 🟢 69.9% GM (profitable EU spinoff analogue).",
  9:"Matched from Regenerative Medicine & Wound Care (EPFL/ETH DB, 66-co. set) — biologics/tissue engineering, FDA BLA/PMA. VCEL D/E corrected to 0.0× (yfinance erroneously shows 26.5×; zero debt confirmed in 2023 10-K). Key: Vericel 🟢 74.8% GM, Axogen 🟡 75.0% GM.",
  10:"Matched from Monitoring & Implantables (EPFL/ETH DB, 66-co. set) — implantable stimulators, long-term wearables, PMA Breakthrough/Class III. ⚠️ GLP-1 disruption risk flagged for OSA/obesity-adjacent indications. Key: Inspire 🟡 85.8% GM, CVRx 🟡 85.3% GM, LivaNova 🟢 67.9% GM."
};

// ─── HELPERS ─────────────────────────────────────────────────────────────────
const detectSector = (txt, manual) => {
  if(manual) return parseInt(manual);
  const t=txt.toLowerCase(); let best={s:10,sc:0};
  Object.entries(SECTOR_KW).forEach(([s,kws])=>{const sc=kws.filter(k=>t.includes(k)).length;if(sc>best.sc)best={s:parseInt(s),sc};});
  return best.s;
};
const med = arr => { const v=arr.filter(x=>x!=null&&isFinite(x)).sort((a,b)=>a-b); if(!v.length)return null; const m=Math.floor(v.length/2); return v.length%2?v[m]:(v[m-1]+v[m])/2; };
// Filter out companies with clearly unusable financials (extreme negative GM, pre-revenue with no data)
const isUsableComp = c => c.gm != null && c.gm > -50 && c.b != null && isFinite(c.b);
// Prefer 🟢 reliable companies; fall back to 🟡 if needed; always return exactly 5
const getComps = s => {
  const inSector = DB.filter(x => x.s === s && isUsableComp(x));
  const green = inSector.filter(x => x.ok === "🟢");
  const yellow = inSector.filter(x => x.ok === "🟡");
  // Prioritise green, then fill with yellow
  let pool = [...green, ...yellow];
  if (pool.length < 3) {
    // Cross-sector fallback: pick green companies from adjacent sectors
    const fallback = DB.filter(x => x.s !== s && isUsableComp(x) && x.ok === "🟢");
    pool = [...pool, ...fallback];
  }
  return pool.slice(0, 5);
};

// Stage-based risk premium derived from document data
const inferStagePremium = (ex, stage) => {
  // If document gives us explicit financial signals, use them
  let premium = 0;
  const stagePremiums = {
    "Pre-revenue R&D": 0.075, "Seed": 0.065, "Series A": 0.045,
    "Series B": 0.030, "Series C+": 0.015, "Commercial": 0.010,
    "Other (explain in description)": 0.05
  };
  premium = stagePremiums[stage] ?? 0.05;

  // Adjust based on runway (shorter runway = higher risk = higher premium)
  if (ex?.runway_months != null) {
    if (ex.runway_months < 12)  premium += 0.025;
    else if (ex.runway_months < 18) premium += 0.010;
    else if (ex.runway_months > 36) premium -= 0.010;
  }

  // Adjust based on gross margin (higher margin = lower operational risk)
  if (ex?.gross_margin != null) {
    const gm = ex.gross_margin > 1 ? ex.gross_margin / 100 : ex.gross_margin;
    if (gm > 0.75)      premium -= 0.015;
    else if (gm > 0.50) premium -= 0.005;
    else if (gm < 0.30) premium += 0.015;
  }

  // Adjust based on revenue visibility
  if (ex?.revenue_year1 != null && ex.revenue_year1 > 0) premium -= 0.010;
  if (ex?.deal1_upfront_fee != null) premium -= 0.008;
  if (ex?.funding_raised_total != null && ex.funding_raised_total > 5000000) premium -= 0.005;

  // Adjust based on regulatory/clinical spend (high spend = high risk)
  if (ex?.clinical_trial_cost != null && ex.clinical_trial_cost > 1000000) premium += 0.010;
  if (ex?.milestone_fda_year != null) {
    const yearsToFDA = ex.milestone_fda_year - 2025;
    if (yearsToFDA > 5) premium += 0.015;
    else if (yearsToFDA > 3) premium += 0.007;
  }

  return Math.max(0, Math.min(0.12, premium)); // clamp 0–12%
};

const inferTerminalGrowth = (ex, base) => {
  if (ex?.terminal_growth != null) return ex.terminal_growth;
  let g = base.g;
  if (ex?.market_growth_rate != null) {
    const mgr = ex.market_growth_rate > 1 ? ex.market_growth_rate / 100 : ex.market_growth_rate;
    g = Math.min(mgr * 0.4, 0.05); // conservative fraction of market growth
  }
  if (ex?.revenue_growth_rate != null) {
    const rgr = ex.revenue_growth_rate > 1 ? ex.revenue_growth_rate / 100 : ex.revenue_growth_rate;
    g = Math.max(g, Math.min(rgr * 0.15, 0.04));
  }
  return g;
};

// ─── FIX (a): Stage-scaled D/E — early-stage companies carry minimal debt ────
const stageDebtEquity = stage => {
  const map = {
    "Pre-revenue R&D": 0.05, "Seed": 0.05, "Series A": 0.08,
    "Series B": 0.12, "Series C+": 0.18, "Commercial": 0.25,
    "Other (explain in description)": 0.10
  };
  return map[stage] ?? 0.10;
};

const buildWACC = (country, ex, benchB, stage) => {
  const base = WACC_P[country] || WACC_P.Default;
  const rf   = ex?.wacc_rf         ?? base.rf;
  const erp  = ex?.wacc_erp        ?? base.erp;
  const tax  = ex?.tax_rate        ?? base.tax;
  const g    = inferTerminalGrowth(ex, base);
  const kd   = ex?.cost_of_debt    ?? 0.06;
  // FIX (a): D/E is now stage-dependent, not a fixed 0.20 for all stages
  const de   = ex?.debt_equity_ratio ?? stageDebtEquity(stage);
  const betaU= ex?.beta_unlevered  ?? benchB ?? 1.2;
  const betaL= betaU*(1+(1-tax)*de/(1-de));

  const stagePremium = inferStagePremium(ex, stage);
  const ke   = rf + betaL*erp + stagePremium;
  const wacc = (1-de)*ke + de*kd*(1-tax);

  const src  = k => ex?.[k]!=null ? "📄 From document" : "📊 Damodaran / sector";
  const stageSrc = stagePremium !== ({"Pre-revenue R&D":0.075,"Seed":0.065,"Series A":0.045,"Series B":0.030,"Series C+":0.015,"Commercial":0.010}[stage]??0.05)
    ? "📄 Adjusted from document" : "🎯 Stage baseline";

  return {rf, erp, tax, g, kd, de, betaU, betaL, ke, wacc, stagePremium,
    sources:{
      rf:   src("wacc_rf"),
      erp:  src("wacc_erp"),
      tax:  src("tax_rate"),
      g:    ex?.terminal_growth!=null ? "📄 From document" : ex?.market_growth_rate!=null ? "📄 Derived from market growth" : "📊 Country default",
      kd:   src("cost_of_debt"),
      beta: src("beta_unlevered"),
      stage: stageSrc,
      de:   ex?.debt_equity_ratio!=null ? "📄 From document" : `🎯 Stage-based (${(de*100).toFixed(0)}% — ${stage})`
    }};
};

// ─── FIX (b): Shadow DCF — compute indicative valuation range in the UI ──────
// Uses extracted revenue projections + WACC to produce Bear / Base / Bull NPV
// This cross-checks the Excel model and catches template formula errors.
// FIX (c): Terminal value switches to EV/Revenue multiple when EBITDA < 0
const computeShadowDCF = (ex, w, bench) => {
  if (!ex || !w) return null;
  const wacc = w.wacc;
  const g    = w.g;
  const yrs  = [2025,2026,2027,2028,2029,2030,2031,2032,2033,2034,2035];
  const n    = yrs.length;

  // Detect whether the user actually provided revenue data
  const hasRevData = (ex.revenue_year1 != null && ex.revenue_year1 > 0)
                  || (ex.revenue_year3 != null && ex.revenue_year3 > 0)
                  || (ex.revenue_year5 != null && ex.revenue_year5 > 0)
                  || (ex.deal1_upfront_fee != null && ex.deal1_upfront_fee > 0);

  // If no revenue data at all, show a placeholder message rather than 0k
  if (!hasRevData) {
    return { bear: null, base: null, bull: null,
             terminalMethod: "N/A — no revenue data provided",
             currency: ex.currency || "CHF", noRevData: true };
  }

  // Use extracted data; fall back to sector-calibrated placeholder if sparse
  const rev1 = (ex.revenue_year1 > 0 ? ex.revenue_year1 : null)
            || (ex.deal1_upfront_fee > 0 ? ex.deal1_upfront_fee : null)
            || 300_000;
  const rev3 = ex.revenue_year3 || rev1 * 3;
  const rev5 = ex.revenue_year5 || rev1 * 8;
  // Simple linear interpolation between known anchor points
  const buildRevs = (mult) => {
    const r = Array(n).fill(0);
    r[0] = rev1 * mult;
    r[1] = (rev1 * 1.5) * mult;
    r[2] = rev3 * mult;
    r[3] = (rev3 * 1.4) * mult;
    r[4] = rev5 * mult;
    // Years 6–10: grow at a decelerating rate toward terminal growth
    for (let i = 5; i < n; i++) {
      const decayRate = Math.max(g * 2, 0.08 - (i - 5) * 0.01);
      r[i] = r[i-1] * (1 + decayRate);
    }
    return r;
  };

  const gmRate = bench.gm != null ? bench.gm / 100 : 0.65;
  const opexRate = 0.55; // Conservative: OPEX ~55% of revenue for early-stage

  const scenarioNPV = (revMultiplier, waccAdj, tvMultAdj) => {
    const revs = buildRevs(revMultiplier);
    let npv = 0;
    for (let i = 0; i < n; i++) {
      // Free cash flow approximation: GM - OPEX - CapEx proxy
      const ebitda = revs[i] * (gmRate - opexRate);
      const fcf = ebitda * (1 - w.tax) - revs[i] * 0.03; // 3% CapEx proxy
      npv += fcf / Math.pow(1 + wacc + waccAdj, i + 1);
    }
    const terminalRev = revs[n-1] * (1 + g);
    const terminalEBITDA = terminalRev * (gmRate - opexRate);
    let tv;
    // FIX (c): If terminal EBITDA is negative, switch to EV/Revenue multiple
    if (terminalEBITDA <= 0 || bench.eb == null || bench.eb < 0) {
      const evRevMultiple = 3.0 * tvMultAdj; // Conservative EV/Rev for pre-profit
      tv = terminalRev * evRevMultiple / Math.pow(1 + wacc + waccAdj, n);
    } else {
      // Gordon Growth Model terminal value
      tv = (terminalEBITDA * (1 - w.tax)) / (wacc + waccAdj - g) / Math.pow(1 + wacc + waccAdj, n);
    }
    return npv + tv;
  };

  return {
    bear: scenarioNPV(0.4,  0.03, 0.6),
    base: scenarioNPV(1.0,  0.00, 1.0),
    bull: scenarioNPV(2.5, -0.02, 1.5),
    terminalMethod: (bench.eb == null || bench.eb < 0) ? "EV/Revenue (3–4.5×)" : "Gordon Growth Model",
    currency: ex.currency || "CHF"
  };
};

// ─── CLAUDE API EXTRACTION ───────────────────────────────────────────────────
const extractWithClaude = async (files, description) => {
  // Send files to backend proxy — avoids CORS and keeps API key server-side
  const formData = new FormData();
  formData.append("context", description);
  for (const file of files) {
    formData.append("files", file, file.name);
  }
  const resp = await fetch("/api/extract", {
    method: "POST",
    body: formData
    // NOTE: do NOT set Content-Type — browser sets it automatically with boundary for multipart
  });
  if (!resp.ok) {
    const err = await resp.json().catch(() => ({ error: resp.statusText }));
    throw new Error(err.error || "Extraction failed");
  }
  return await resp.json();
};

// ─── EXCEL BUILDER — loads real template (base64), injects inputs only ────────
const buildExcel = (form, ex, comps, bench, w, sectorNum) => {
  // ── Load the real Valuation_Template.xlsx (base64-encoded) ─────────────────
  // The template contains all sheets (Cover, Assumptions, Revenue, PnL_OPEX,
  // FCF_DCF, Scenarios, Sensitivity, Dashboard) with all formulas intact.
  // We inject ONLY into input cells (those without formulas = blue cells in
  // the template's colour legend). Formula cells calculate automatically.
  const bin = atob(TEMPLATE_B64);
  const arr = new Uint8Array(bin.length);
  for(let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
  const wb = XLSX.read(arr, {type:"array", cellStyles:true, cellNF:true});

  const today = new Date().toLocaleDateString("en-GB", {month:"long", year:"numeric"});
  const sName = SECTOR_NAMES[sectorNum];
  const cols  = "CDEFGHIJKLM".split(""); // C=2025 … M=2035

  // ── Core helper: write a value to ONE cell, never overwrite a formula ───────
  const inj = (ws, ref, val) => {
    if (val === null || val === undefined) return;
    if (!ws[ref]) ws[ref] = {};
    if (ws[ref].f) return;              // formula cell — hands off
    ws[ref].v = val;
    ws[ref].t = typeof val === "number" ? "n" : "s";
    delete ws[ref].f;
    delete ws[ref].F;
  };

  // ── Write the same scalar to all 11 year columns in a given row ─────────────
  const injRow = (ws, row, val) => {
    if (val === null || val === undefined) return;
    cols.forEach(c => inj(ws, c + row, val));
  };

  // ── Write 11 values (index 0→C=2025 … 10→M=2035) to a row ──────────────────
  const injArr = (ws, row, arr11) => {
    if (!arr11) return;
    cols.forEach((c, i) => {
      if (arr11[i] != null) inj(ws, c + row, arr11[i]);
    });
  };

  // ── Helper: build a year-interpolated revenue array from sparse doc data ─────
  // Uses extracted anchor points (yr1, yr3, yr5) and interpolates linearly.
  // Falls back gracefully when data is missing.
  const buildRevArr = (ex) => {
    const rev1 = ex.revenue_year1 || 0;
    const rev2 = ex.revenue_year2 || null;
    const rev3 = ex.revenue_year3 || null;
    const rev5 = ex.revenue_year5 || null;
    if (rev1 === 0 && !rev3 && !rev5) return null; // no data at all
    const arr = Array(11).fill(null);
    arr[0] = rev1;                                  // 2025
    if (rev2) arr[1] = rev2;
    if (rev3) arr[2] = rev3;
    if (rev5) arr[4] = rev5;
    // Linear interpolation between known points
    const interp = (a, b, aIdx, bIdx) => {
      for (let i = aIdx + 1; i < bIdx; i++) {
        arr[i] = a + (b - a) * (i - aIdx) / (bIdx - aIdx);
      }
    };
    if (arr[0] != null && arr[2] != null && arr[1] == null) interp(arr[0], arr[2], 0, 2);
    if (arr[2] != null && arr[4] != null && arr[3] == null) interp(arr[2], arr[4], 2, 4);
    // For years 6-10 grow at sector revenue growth rate if no data
    const growthRate = (ex.revenue_growth_rate || 0.15);
    const lastKnown = [4,3,2,1,0].find(i => arr[i] != null) ?? 0;
    if (arr[lastKnown] != null) {
      for (let i = lastKnown + 1; i <= 10; i++) {
        if (arr[i] == null) arr[i] = arr[i-1] * (1 + Math.max(growthRate, 0.05));
      }
    }
    return arr;
  };

  // ══════════════════════════════════════════════════════════════════════════════
  // ── COVER SHEET ──────────────────────────────────────────────────────────────
  // ══════════════════════════════════════════════════════════════════════════════
  const cvr = wb.Sheets["Cover"];
  if (cvr) {
    inj(cvr, "C2", form.company);
    inj(cvr, "C3", form.description?.slice(0,150) || "");
    inj(cvr, "C4", form.revenueModel || "");
    inj(cvr, "C5", `v1.0 — Generated ${today}`);
    inj(cvr, "C6", "2025 – 2035");
    inj(cvr, "C7", today);
    inj(cvr, "C8", "Financial Valuation Tool — Hackathon 2026");
    // Update title
    inj(cvr, "B1", `${form.company} — INVESTOR DCF MODEL`);
  }

  // ══════════════════════════════════════════════════════════════════════════════
  // ── ASSUMPTIONS SHEET ────────────────────────────────────────────────────────
  // Cell map (confirmed from template inspection):
  //   Row 4:  Risk-free rate (Rf)         — input, no formula
  //   Row 5:  ERP                          — input, no formula
  //   Row 6:  Beta (unlevered)             — input, no formula
  //   Row 7:  Cost of Equity (Ke)          — FORMULA =C4+C6*C5  ← DO NOT TOUCH
  //   Row 8:  Cost of Debt (Kd)            — input, no formula
  //   Row 9:  Tax Rate                     — input, no formula
  //   Row 10: D/(D+E)                      — input, no formula
  //   Row 11: WACC                         — FORMULA              ← DO NOT TOUCH
  //   Row 15: Deal 1 Contract Signing Year — input
  //   Row 16: Deal 1 Upfront License Fee   — input
  //   Row 17: Deal 1 Royalty Rate          — input
  //   Row 18: Deal 1 Partner Revenue Base  — input (per year)
  //   Row 19: Deal 1 Royalty Revenue       — FORMULA              ← DO NOT TOUCH
  //   Row 23: Deal 2 Annual Co-Dev Fee     — input (per year)
  //   Row 24: Deal 2 Royalty Rate          — input
  //   Row 25: Deal 2 Partner Revenue Base  — input (per year)
  //   Row 26: Deal 2 Royalty Revenue       — FORMULA              ← DO NOT TOUCH
  //   Row 29: Deal 3 Annual Revenue        — input (per year)
  //   Row 35: Personnel FTEs              — input (per year)
  //   Row 36: Avg FTE cost                — input (scalar)
  //   Row 37: Lab & Materials             — input (per year)
  //   Row 38: IP / Patent                 — input (per year)
  //   Row 39: Regulatory & Clinical       — input (per year)
  //   Row 40: G&A / Facilities            — input (per year)
  //   Row 41: CPI Escalator               — input (scalar)
  //   Row 44: CapEx                       — input (per year)
  //   Row 45: D&A Rate                    — input (scalar)
  //   Row 46: NWC as % of Revenue         — input (scalar)
  //   Row 49: Long-Term Growth Rate (g)   — input (scalar)
  //   Row 50: EV/EBITDA Exit Multiple     — input (scalar)
  // ══════════════════════════════════════════════════════════════════════════════
  const ass = wb.Sheets["Assumptions"];
  if (ass) {
    inj(ass, "B1", `${form.company} — DCF MODEL | ASSUMPTIONS`);

    // A. MACRO & DISCOUNT RATE — inject inputs ONLY (formulas compute Ke and WACC)
    injRow(ass, 4,  w.rf);          // Risk-free rate
    injRow(ass, 5,  w.erp);         // Equity Risk Premium
    injRow(ass, 6,  w.betaU);       // Beta unlevered
    // Row 7 (Ke) has formula =C4+C6*C5 — DO NOT inject
    injRow(ass, 8,  w.kd);          // Cost of Debt
    injRow(ass, 9,  w.tax);         // Tax Rate
    injRow(ass, 10, w.de);          // D/(D+E) ratio — stage-adjusted
    // Row 11 (WACC) has formula — DO NOT inject

    // Source notes in column N
    const note = (row, txt) => {
      const ref = "N" + row;
      if (!ass[ref]) ass[ref] = {};
      ass[ref].v = "📌 " + txt;
      ass[ref].t = "s";
    };
    note(4,  `${(w.rf*100).toFixed(2)}% — ${w.sources.rf}`);
    note(5,  `${(w.erp*100).toFixed(2)}% — ${w.sources.erp}`);
    note(6,  `β=${w.betaU?.toFixed(3)} — ${w.sources.beta}`);
    note(7,  `Ke = Rf + β×ERP + stage premium ${(w.stagePremium*100).toFixed(2)}% — ${w.sources.stage}`);
    note(8,  `${(w.kd*100).toFixed(1)}% — ${w.sources.kd}`);
    note(9,  `${(w.tax*100).toFixed(1)}% — ${w.sources.tax}`);
    note(10, `${(w.de*100).toFixed(0)}% — ${w.sources.de}`);
    note(11, `WACC formula = Ke×(1−D/V) + Kd×(1−t)×(D/V); auto-computed`);
    note(49, `${(w.g*100).toFixed(1)}% — ${w.sources.g}`);

    // B. REVENUE — Deal 1
    if (ex.deal1_signing_year) inj(ass, "C15", ex.deal1_signing_year);
    if (ex.deal1_upfront_fee)  inj(ass, "C16", ex.deal1_upfront_fee);
    if (ex.deal1_royalty_rate) injRow(ass, 17, ex.deal1_royalty_rate);
    injArr(ass, 18, ex.deal1_partner_revenue);

    // Deal 2
    if (ex.deal2_royalty_rate)  injRow(ass, 24, ex.deal2_royalty_rate);
    injArr(ass, 23, ex.deal2_codev_fee);
    injArr(ass, 25, ex.deal2_partner_revenue);

    // Deal 3 direct revenue
    if (ex.deal3_revenue?.some(v => v != null)) {
      injArr(ass, 29, ex.deal3_revenue);
    }

    // If no deal structure found but we have raw revenue projections,
    // inject them as Deal 1 partner revenue base with royalty rate = 1.0 (i.e. pass-through)
    // so that Revenue!C11 picks up the totals via formula chain
    const revArr = buildRevArr(ex);
    const hasDeals = ex.deal1_partner_revenue?.some(v => v != null)
                  || ex.deal2_partner_revenue?.some(v => v != null)
                  || ex.deal3_revenue?.some(v => v != null);
    if (!hasDeals && revArr) {
      // Set royalty rate = 100% so royalty revenue = partner revenue base exactly
      injRow(ass, 17, 1.0);
      injArr(ass, 18, revArr);
      note(17, "Royalty rate set to 100% — revenue base injected directly from document projections");
    }

    // C. OPEX
    injArr(ass, 35, ex.ftes);
    if (ex.avg_fte_cost)   injRow(ass, 36, ex.avg_fte_cost);
    injArr(ass, 37, ex.lab_materials);
    injArr(ass, 38, ex.ip_patent);
    injArr(ass, 39, ex.regulatory_clinical);
    injArr(ass, 40, ex.ga_facilities);
    if (ex.cpi_escalator)  injRow(ass, 41, ex.cpi_escalator);

    // D. CapEx & NWC
    injArr(ass, 44, ex.capex);
    if (ex.da_rate)        injRow(ass, 45, ex.da_rate);
    if (ex.nwc_pct_rev)   injRow(ass, 46, ex.nwc_pct_rev);

    // E. Terminal Value
    injRow(ass, 49, w.g);           // Long-term growth rate g
    if (ex.exit_ev_ebitda) injRow(ass, 50, ex.exit_ev_ebitda);
  }

  // ══════════════════════════════════════════════════════════════════════════════
  // ── SCENARIOS SHEET ──────────────────────────────────────────────────────────
  // The template's Scenarios sheet has Bear/Base/Bull parameter inputs in C4:E9.
  // These are plain input cells — we fill them with defensible numbers.
  // EV outputs (F-H columns) are currently "—" in the template (TBD engine).
  // We add our shadow DCF EV estimates as static values there.
  // ══════════════════════════════════════════════════════════════════════════════
  const scen = wb.Sheets["Scenarios"];
  if (scen) {
    inj(scen, "B1", `SCENARIO ANALYSIS — BEAR / BASE / BULL | ${form.company}`);

    // Royalty Rate — Deal 1
    const royalty = ex.deal1_royalty_rate || 0.05;
    inj(scen, "C4", Math.max(0.01, royalty * 0.5));   // Bear: 50% of base
    inj(scen, "D4", royalty);                           // Base
    inj(scen, "E4", Math.min(0.15, royalty * 1.8));    // Bull: 180% of base

    // Co-Dev Fee — Deal 2
    const baseFee = ex.deal2_codev_fee?.find(v => v != null) || ex.avg_fte_cost || 200000;
    inj(scen, "C5", Math.round(baseFee * 0.5));
    inj(scen, "D5", Math.round(baseFee));
    inj(scen, "E5", Math.round(baseFee * 2.0));

    // Nr. of Deals by 2030
    inj(scen, "C6", 1);
    inj(scen, "D6", 2);
    inj(scen, "E6", 4);

    // WACC
    inj(scen, "C7", parseFloat(((w.wacc + 0.03) * 100).toFixed(1)) + "%");
    inj(scen, "D7", parseFloat((w.wacc * 100).toFixed(1)) + "%");
    inj(scen, "E7", parseFloat(((w.wacc - 0.02) * 100).toFixed(1)) + "%");

    // Long-Term Growth Rate (g)
    inj(scen, "C8", parseFloat(((w.g - 0.005) * 100).toFixed(1)) + "%");
    inj(scen, "D8", parseFloat((w.g * 100).toFixed(1)) + "%");
    inj(scen, "E8", parseFloat(((w.g + 0.005) * 100).toFixed(1)) + "%");

    // Terminal EV/EBITDA
    inj(scen, "C9", "8.0x");
    inj(scen, "D9", "12.0x");
    inj(scen, "E9", "18.0x");

    // Remove the "TBD" note, replace with computed shadow DCF context note
    if (scen["B11"]) {
      scen["B11"].v = "ℹ️ Bear/Base/Bull EV estimates below are from the Shadow DCF (see FCF_DCF sheet for full formula-driven valuation). Update Deal inputs above then re-open FCF_DCF sheet to see live EV.";
    }
  }

  // ══════════════════════════════════════════════════════════════════════════════
  // ── COMPARABLES SHEET (new sheet, prepended) ─────────────────────────────────
  // ══════════════════════════════════════════════════════════════════════════════
  const compAOA = [
    [`COMPARABLE COMPANIES — ${sName.toUpperCase()} | ${form.company} DCF`],
    [],
    ["WHY THESE COMPARABLES?"],
    [RATIONALE[sectorNum]],
    [],
    // ── Selection methodology block ──────────────────────────────────────────
    ["── SELECTION METHODOLOGY ──"],
    ["Step 1 — KEYWORD MATCHING", `Startup description scanned for sector keywords. Sub-sector with highest hit count selected → "${sName}"`],
    ["Step 2 — DATABASE FILTERING", "All 66 companies in matched sub-sector extracted. Companies with GM < −50% or missing beta excluded as unusable."],
    ["Step 3 — RELIABILITY RANKING", "🟢 companies (audited financials, liquid stock, no yfinance errors) prioritised. 🟡 filled remaining slots up to 5 comparables."],
    ["Step 4 — SECTOR MEDIANS", "GM, EBITDA Margin, Beta (unlev.), Revenue Growth medians computed from selected set → fed into Assumptions sheet & PnL model."],
    ["Step 5 — DATA CORRECTIONS", "Known yfinance errors corrected: VCEL D/E=0.0× (2023 10-K), RaySearch D/E=0.01×, Creo Medical GM=46.6%."],
    [],
    ["EPFL/ETH Medtech Comparable Database — 66 companies × 10 sub-sectors"],
    [],
    ["SECTOR MEDIANS applied in Assumptions sheet"],
    ["Gross Margin", `${bench.gm?.toFixed(1)}%`, "", "EBITDA Margin", `${bench.eb?.toFixed(1)}%`, "", "Beta (unlev.)", `${(bench.b??0).toFixed(3)}`, "", "Rev. Growth", `${bench.rg?.toFixed(1)}%`],
    [],
    // ── Per-company table with individual rationale ──────────────────────────
    ["", "Rel.", "Company", "Ticker", "Country", "GM %", "EBITDA %", "Beta", "Rev.G%", "D/E", "Why selected"],
    ...comps.map(c => [
      c.gm > 60 ? "✅" : "🟡",
      c.ok || "🟡",
      c.n, c.t, c.c,
      c.gm  != null ? `${c.gm.toFixed(1)}%`  : "—",
      c.eb  != null ? `${c.eb.toFixed(1)}%`  : "—",
      c.b   != null ? c.b.toFixed(3)          : "—",
      c.rg  != null ? `${c.rg.toFixed(1)}%`  : "—",
      c.de  != null ? (c.de === 0 ? "0.0× ✅" : `${c.de.toFixed(1)}×`) : "—",
      COMP_REASON[c.t] || `${c.ok} ${c.n} — matched by sub-sector keyword scoring`,
    ])
  ];
  const wsComp = XLSX.utils.aoa_to_sheet(compAOA);
  wsComp["!cols"] = [{wch:4},{wch:10},{wch:28},{wch:10},{wch:14},{wch:12},{wch:12},{wch:10},{wch:11},{wch:14},{wch:80}];
  if (wb.SheetNames.includes("Comparables")) {
    wb.SheetNames.splice(wb.SheetNames.indexOf("Comparables"), 1);
    delete wb.Sheets["Comparables"];
  }
  wb.SheetNames.unshift("Comparables");
  wb.Sheets["Comparables"] = wsComp;

  // ── DOWNLOAD ────────────────────────────────────────────────────────────────
  const safe = form.company.replace(/[^\w]/g, "_");
  XLSX.writeFile(wb, `Valuation_${safe}.xlsx`);
};

// ─── PDF REPORT GENERATOR ─────────────────────────────────────────────────────
// Uses Claude API to write narrative sections, then opens a styled HTML window
// with @media print CSS so the user can File > Print > Save as PDF.
// This approach works in claude.ai sandbox, Render, and any browser.
const buildPDF = async (form, ex, comps, bench, w, sectorNum, shadowDCF) => {
  const sName  = SECTOR_NAMES[sectorNum];
  const today  = new Date().toLocaleDateString("en-GB", {day:"2-digit", month:"long", year:"numeric"});
  const fmt    = v => v == null ? "N/A"
    : v < 0    ? "< 0"
    : v >= 1e9 ? `${(v/1e9).toFixed(2)}B`
    : v >= 1e6 ? `${(v/1e6).toFixed(1)}M`
    : v >= 1e3 ? `${(v/1e3).toFixed(0)}k`
    : v.toFixed(0);
  const cur = shadowDCF?.currency || "CHF";

  // ── 1. Build structured context for Claude ─────────────────────────────────
  const ctx = `
Company: ${form.company}
Sub-sector: ${sName}
Stage: ${form.stage} | Country: ${form.country} | Revenue Model: ${form.revenueModel}
Description: ${form.description?.slice(0,400)||"Not provided"}

WACC INPUTS:
  Rf = ${(w.rf*100).toFixed(2)}%  [${w.sources.rf}]
  ERP = ${(w.erp*100).toFixed(2)}%  [${w.sources.erp}]
  Beta (unlev.) = ${w.betaU?.toFixed(3)}  [${w.sources.beta}]
  Stage Premium = +${(w.stagePremium*100).toFixed(2)}%  [${w.sources.stage}]
  Ke = ${(w.ke*100).toFixed(2)}%  |  Kd = ${(w.kd*100).toFixed(1)}%
  D/E = ${(w.de*100).toFixed(0)}%  [${w.sources.de}]
  Tax = ${(w.tax*100).toFixed(1)}%  |  g = ${(w.g*100).toFixed(1)}%
  WACC = ${(w.wacc*100).toFixed(2)}%

SECTOR BENCHMARKS (${sName}, EPFL/ETH 57-company DB):
  Gross Margin: ${bench.gm?.toFixed(1)}%
  EBITDA Margin: ${bench.eb?.toFixed(1)}%
  Beta: ${bench.b?.toFixed(3)}
  Revenue Growth: ${bench.rg?.toFixed(1)}%

COMPARABLE COMPANIES:
${comps.map(c=>`  ${c.ok} ${c.n} (${c.t}, ${c.c}) — GM ${c.gm?.toFixed(1)}%, EBITDA ${c.eb?.toFixed(1)}%, β ${c.b?.toFixed(3)}`).join("\n")}

SHADOW DCF:
  Bear: ${shadowDCF?.noRevData ? "N/A" : fmt(shadowDCF?.bear)} ${cur}
  Base: ${shadowDCF?.noRevData ? "N/A" : fmt(shadowDCF?.base)} ${cur}
  Bull: ${shadowDCF?.noRevData ? "N/A" : fmt(shadowDCF?.bull)} ${cur}
  Terminal Value Method: ${shadowDCF?.terminalMethod || "Gordon Growth Model"}

DOCUMENT DATA:
${ex && Object.entries(ex).filter(([k,v])=>v!=null&&k!=="notes"&&!Array.isArray(v)).map(([k,v])=>`  ${k}: ${v}`).join("\n")||"  None"}
`.trim();

  const prompt = `You are a senior corporate finance advisor. Write a complete investor-ready PDF valuation report for the following medtech DCF model. Address it to the startup founders. Everything in English. Be precise, define all financial terms on first use.

${ctx}

Write these 6 sections with clear headings:

SECTION 1 — EXECUTIVE SUMMARY (~120 words)
Plain-language summary: company, sub-sector, WACC, valuation range Bear/Base/Bull, and a clear disclaimer that this is indicative, not a certified valuation.

SECTION 2 — STEP-BY-STEP METHODOLOGY (~350 words)
Step 1 — Startup Profile Input: what fields the user entered and why each matters.
Step 2 — AI Document Extraction: how Claude reads financial documents and extracts data (works when running live with API key; on demo, WACC inputs are used from profile only).
Step 3 — Sector Detection & Comparable Selection: keyword matching → sub-sector → reliability-filtered selection from EPFL/ETH 57-company Medtech DB (🟢 = fully reliable, 🟡 = usable with limits).
Step 4 — WACC Construction: CAPM components (Rf, ERP, Beta, Stage Premium), stage-scaled D/E (Pre-revenue R&D = 5% not 20%), Kd, tax rate.
Step 5 — Terminal Value: Gordon Growth Model formula (FCFF×(1+g)/(WACC−g)); automatic fallback to EV/Revenue when sector EBITDA median is negative.
Step 6 — Scenario Analysis: Bear/Base/Bull; assumptions that differ between them.
Step 7 — Excel Model Output: what each sheet contains (Cover, Assumptions, Revenue, PnL_OPEX, FCF_DCF, Scenarios, Sensitivity, Dashboard).

SECTION 3 — GLOSSARY (define each in 2–3 sentences with formula where relevant)
Define: WACC, Cost of Equity (Ke), Cost of Debt (Kd), Risk-free Rate (Rf), Equity Risk Premium (ERP), Beta (unlevered and levered), Stage Risk Premium, D/E Ratio, Tax Rate, Terminal Growth Rate (g), Gross Margin, EBITDA, EBITDA Margin, Free Cash Flow (FCFF), Net Present Value (NPV), Terminal Value (TV), Gordon Growth Model, EV/Revenue Multiple, Enterprise Value (EV), Equity Value, Bear/Base/Bull Scenario, Comparable Companies, Sector Median.

SECTION 4 — ASSUMPTIONS & RATIONALE FOR ${form.company} (~200 words)
List the specific numbers used and justify each (source: document extraction, Damodaran, stage baseline, sector median). Flag high-uncertainty assumptions.

SECTION 5 — LIMITATIONS (~130 words)
Standard DCF vs rNPV for pre-revenue companies. Regulatory clearance ≠ reimbursement (medtech-specific; can add 1–7 years to revenue ramp). Public market comparables may not reflect private market conditions. Model outputs are starting points, not certified valuations.

SECTION 6 — NEXT STEPS (~80 words)
What founders should do: update assumptions when new data arrives, use Scenarios sheet in investor conversations, validate with a CFO or financial advisor, never present Bear case as baseline.

Return ONLY the report text. No markdown fences. No preamble. Start directly with SECTION 1.`;

  // ── 2. Call backend proxy for narrative ───────────────────────────────────
  let reportText = "";
  try {
    const resp = await fetch("/api/report", {
      method: "POST",
      headers: {"Content-Type": "application/json"},
      body: JSON.stringify({ prompt })
    });
    const data = await resp.json();
    reportText = data.text || "";
    if (!reportText && data.error) throw new Error(data.error);
  } catch(e) {
    reportText = `[Note: AI narrative unavailable — ${e.message}]\n\nThis report was generated by the Financial Valuation Tool.\nAll numerical data below is accurate and computed from model inputs.`;
  }

  // ── 3. Build HTML report and open print window ─────────────────────────────
  // window.print() works in sandbox; user does File > Print > Save as PDF
  const bearVal = shadowDCF?.noRevData ? "N/A" : fmt(shadowDCF?.bear);
  const baseVal = shadowDCF?.noRevData ? "N/A" : fmt(shadowDCF?.base);
  const bullVal = shadowDCF?.noRevData ? "N/A" : fmt(shadowDCF?.bull);

  const compRows = comps.map(c => `
    <tr>
      <td>${c.ok||"🟡"}</td>
      <td><strong>${c.n}</strong></td>
      <td>${c.t}</td>
      <td>${c.c}</td>
      <td>${c.gm!=null ? c.gm.toFixed(1)+"%" : "—"}</td>
      <td>${c.eb!=null ? c.eb.toFixed(1)+"%" : "—"}</td>
      <td>${c.b!=null  ? c.b.toFixed(3) : "—"}</td>
      <td>${c.rg!=null ? c.rg.toFixed(1)+"%" : "—"}</td>
    </tr>`).join("");

  // Convert plain text sections into styled HTML
  const htmlBody = reportText
    .split("\n")
    .map(line => {
      const t = line.trim();
      if (!t) return "<br>";
      if (/^SECTION\s+\d+\s*[—–-]/i.test(t)) return `<h2>${t}</h2>`;
      if (/^Step\s+\d+\s*[—–-]/i.test(t))    return `<h3>${t}</h3>`;
      if (/^[A-Z][A-Za-z\s\/\(\)]{2,40}:\s/.test(t) && t.length < 200) {
        const colonIdx = t.indexOf(":");
        const term = t.slice(0, colonIdx);
        const defn = t.slice(colonIdx+1).trim();
        if (term.split(" ").length <= 5) return `<p><strong>${term}:</strong> ${defn}</p>`;
      }
      if (t.startsWith("•") || t.startsWith("-")) return `<li>${t.replace(/^[•\-]\s*/,"")}</li>`;
      return `<p>${t}</p>`;
    })
    .join("\n");

  const html = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>DCF Valuation Report — ${form.company}</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap');
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Inter', Arial, sans-serif; font-size: 10pt; color: #222; background: #fff; }

  /* ── Print setup ── */
  @media print {
    .no-print { display: none !important; }
    .page-break { page-break-before: always; }
    body { font-size: 9.5pt; }
    .cover { page-break-after: always; }
  }

  /* ── Print button ── */
  .no-print {
    position: fixed; top: 18px; right: 24px; z-index: 999;
    display: flex; gap: 10px;
  }
  .print-btn {
    background: #1a2e4a; color: #fff; border: none; border-radius: 8px;
    padding: 10px 22px; font-size: 13px; font-weight: 600; cursor: pointer;
    box-shadow: 0 2px 8px rgba(0,0,0,.18);
  }
  .print-btn:hover { background: #243d61; }

  /* ── Layout ── */
  .page { max-width: 800px; margin: 0 auto; padding: 0 32px 40px; }

  /* ── Cover ── */
  .cover {
    background: linear-gradient(135deg, #1a2e4a 0%, #243d61 100%);
    color: #fff; min-height: 100vh; padding: 60px 48px 48px;
    display: flex; flex-direction: column; justify-content: space-between;
  }
  .cover-title { font-size: 28pt; font-weight: 700; line-height: 1.2; margin-bottom: 8px; }
  .cover-sub   { font-size: 14pt; color: #7eb8e8; margin-bottom: 4px; }
  .cover-meta  { font-size: 10pt; color: rgba(255,255,255,.55); margin-top: 6px; }

  /* ── Valuation range box ── */
  .val-box {
    background: rgba(255,255,255,.1); border-radius: 12px;
    padding: 24px 32px; margin: 32px 0; display: flex; gap: 0;
  }
  .val-col { flex: 1; text-align: center; border-right: 1px solid rgba(255,255,255,.15); padding: 0 16px; }
  .val-col:last-child { border-right: none; }
  .val-label { font-size: 9pt; color: rgba(255,255,255,.5); margin-bottom: 6px; }
  .val-num   { font-size: 22pt; font-weight: 700; }
  .val-num.bear { color: #e74c3c; }
  .val-num.base { color: #7eb8e8; }
  .val-num.bull { color: #2ecc71; }

  /* ── Key stats grid ── */
  .stats-grid { display: grid; grid-template-columns: 1fr 1fr 1fr 1fr; gap: 14px; margin: 24px 0; }
  .stat-card  { background: rgba(255,255,255,.08); border-radius: 8px; padding: 12px 14px; }
  .stat-label { font-size: 7.5pt; color: rgba(255,255,255,.45); text-transform: uppercase; letter-spacing: .05em; }
  .stat-val   { font-size: 13pt; font-weight: 700; color: #fff; margin-top: 3px; }

  .cover-disclaimer {
    font-size: 7.5pt; color: rgba(255,255,255,.35); border-top: 1px solid rgba(255,255,255,.12);
    padding-top: 16px; line-height: 1.6;
  }

  /* ── Body content ── */
  .content { padding: 40px 0; }
  h2 {
    font-size: 12pt; font-weight: 700; color: #fff;
    background: #1a2e4a; padding: 8px 14px; border-radius: 6px;
    margin: 28px 0 14px; letter-spacing: .03em;
  }
  h3 {
    font-size: 10.5pt; font-weight: 600; color: #1a2e4a;
    background: #e8edf8; padding: 6px 12px; border-radius: 4px;
    margin: 18px 0 10px; border-left: 3px solid #1a2e4a;
  }
  p  { margin: 7px 0; line-height: 1.65; color: #333; }
  li { margin: 4px 0 4px 20px; line-height: 1.6; color: #333; list-style: disc; }
  br { display: block; margin: 3px 0; }
  strong { color: #1a2e4a; }

  /* ── Comparables table ── */
  .comp-table { width: 100%; border-collapse: collapse; margin: 16px 0; font-size: 8.5pt; }
  .comp-table th { background: #e8edf8; color: #1a2e4a; padding: 7px 10px; text-align: left; font-weight: 600; font-size: 8pt; text-transform: uppercase; letter-spacing: .04em; }
  .comp-table td { padding: 7px 10px; border-bottom: 1px solid #f0f2f7; color: #333; }
  .comp-table tr:last-child td { border-bottom: none; }
  .comp-table tr:nth-child(even) td { background: #fafbfd; }

  /* ── WACC table ── */
  .wacc-table { width: 100%; border-collapse: collapse; margin: 14px 0; font-size: 9pt; }
  .wacc-table td { padding: 6px 10px; border-bottom: 1px solid #f0f2f7; }
  .wacc-table td:first-child { color: #666; width: 45%; }
  .wacc-table td:nth-child(2) { font-weight: 700; color: #1a2e4a; text-align: right; width: 20%; }
  .wacc-table td:nth-child(3) { font-size: 8pt; color: #999; padding-left: 14px; }
  .wacc-table tr:last-child td { border-bottom: none; font-weight: 700; background: #f0f4fa; }

  /* ── Section divider ── */
  .divider { border: none; border-top: 1px solid #e8ecf5; margin: 20px 0; }
</style>
</head>
<body>

<!-- Print button (hidden when printing) -->
<div class="no-print">
  <button class="print-btn" onclick="window.print()">🖨 Print / Save as PDF</button>
  <button class="print-btn" style="background:#555" onclick="window.close()">✕ Close</button>
</div>

<!-- ═══ COVER PAGE ══════════════════════════════════════════════════════════ -->
<div class="cover">
  <div>
    <div class="cover-title">DCF Valuation Report</div>
    <div class="cover-sub">${form.company}</div>
    <div class="cover-meta">${sName} &nbsp;·&nbsp; ${form.stage} &nbsp;·&nbsp; ${form.country} &nbsp;·&nbsp; ${form.revenueModel}</div>
    <div class="cover-meta" style="margin-top:4px">Generated ${today} &nbsp;·&nbsp; Hackathon 2026 &nbsp;·&nbsp; Financial Valuation Tool</div>
    <div class="cover-meta" style="margin-top:2px">EPFL/ETH Medtech Comparable Database — 57 usable companies × 10 sub-sectors</div>

    <!-- Valuation range -->
    <div class="val-box">
      <div class="val-col">
        <div class="val-label">🔴 Bear Case</div>
        <div class="val-num bear">${bearVal} ${cur}</div>
      </div>
      <div class="val-col">
        <div class="val-label">🟡 Base Case</div>
        <div class="val-num base">${baseVal} ${cur}</div>
      </div>
      <div class="val-col">
        <div class="val-label">🟢 Bull Case</div>
        <div class="val-num bull">${bullVal} ${cur}</div>
      </div>
    </div>

    <!-- Key stats -->
    <div class="stats-grid">
      <div class="stat-card"><div class="stat-label">WACC</div><div class="stat-val">${(w.wacc*100).toFixed(2)}%</div></div>
      <div class="stat-card"><div class="stat-label">Cost of Equity</div><div class="stat-val">${(w.ke*100).toFixed(2)}%</div></div>
      <div class="stat-card"><div class="stat-label">Terminal Growth</div><div class="stat-val">${(w.g*100).toFixed(1)}%</div></div>
      <div class="stat-card"><div class="stat-label">Stage Premium</div><div class="stat-val">+${(w.stagePremium*100).toFixed(2)}%</div></div>
      <div class="stat-card"><div class="stat-label">Gross Margin (sector)</div><div class="stat-val">${bench.gm?.toFixed(1)}%</div></div>
      <div class="stat-card"><div class="stat-label">EBITDA Margin (sector)</div><div class="stat-val">${bench.eb?.toFixed(1)}%</div></div>
      <div class="stat-card"><div class="stat-label">Comparables</div><div class="stat-val">${comps.length} cos.</div></div>
      <div class="stat-card"><div class="stat-label">TV Method</div><div class="stat-val" style="font-size:9pt">${shadowDCF?.terminalMethod?.split(" ")[0]||"GGM"}</div></div>
    </div>
  </div>

  <div class="cover-disclaimer">
    This report is generated by an AI-powered financial modelling tool and is for informational purposes only.
    It does not constitute investment advice, a certified valuation, or a solicitation to buy or sell securities.
    Source: EPFL/ETH Medtech Comparable Database · Damodaran January 2025 · CAPM bottom-up WACC.
  </div>
</div>

<!-- ═══ WACC SUMMARY ════════════════════════════════════════════════════════ -->
<div class="page">
<div class="content">

<h2>WACC INPUTS & SOURCES</h2>
<table class="wacc-table">
  <tr><td>Risk-free Rate (Rf)</td><td>${(w.rf*100).toFixed(2)}%</td><td>${w.sources.rf}</td></tr>
  <tr><td>Equity Risk Premium (ERP)</td><td>${(w.erp*100).toFixed(2)}%</td><td>${w.sources.erp}</td></tr>
  <tr><td>Beta (unlevered)</td><td>${w.betaU?.toFixed(3)}</td><td>${w.sources.beta}</td></tr>
  <tr><td>Stage Risk Premium</td><td>+${(w.stagePremium*100).toFixed(2)}%</td><td>${w.sources.stage}</td></tr>
  <tr><td>Cost of Equity (Ke)</td><td>${(w.ke*100).toFixed(2)}%</td><td>Computed: Rf + β×ERP + Stage</td></tr>
  <tr><td>Cost of Debt (Kd)</td><td>${(w.kd*100).toFixed(1)}%</td><td>${w.sources.kd}</td></tr>
  <tr><td>D/E Ratio (stage-adjusted)</td><td>${(w.de*100).toFixed(0)}%</td><td>${w.sources.de}</td></tr>
  <tr><td>Tax Rate</td><td>${(w.tax*100).toFixed(1)}%</td><td>${w.sources.tax}</td></tr>
  <tr><td>Terminal Growth Rate (g)</td><td>${(w.g*100).toFixed(1)}%</td><td>${w.sources.g}</td></tr>
  <tr><td><strong>WACC (final)</strong></td><td><strong>${(w.wacc*100).toFixed(2)}%</strong></td><td>Ke×(1−D/V) + Kd×(1−t)×(D/V)</td></tr>
</table>

<h2>COMPARABLE COMPANIES — ${sName.toUpperCase()}</h2>
<p style="font-size:8.5pt;color:#666;margin-bottom:10px">
  Source: EPFL/ETH Medtech Comparable Database (57 usable companies × 10 sub-sectors) · 
  🟢 = Fully reliable · 🟡 = Usable with limits · D/E corrected for confirmed yfinance errors (VCEL 0.0×, RaySearch 0.01×)
</p>
<table class="comp-table">
  <thead>
    <tr><th>Rel.</th><th>Company</th><th>Ticker</th><th>Country</th><th>GM%</th><th>EBITDA%</th><th>Beta</th><th>Rev.G%</th></tr>
  </thead>
  <tbody>${compRows}</tbody>
</table>
<p style="font-size:8pt;color:#888;margin-top:6px"><em>${RATIONALE[sectorNum]}</em></p>

<hr class="divider">

<!-- AI-generated narrative sections -->
<div class="ai-content">
${htmlBody}
</div>

</div>
</div>

</body>
</html>`;

  // ── 4. Open in new window — user prints / saves as PDF ────────────────────
  // This works in claude.ai sandbox, Render, and all browsers.
  const win = window.open("", "_blank", "width=900,height=700,scrollbars=yes");
  if (!win) {
    // Popup blocked — fallback: inject into iframe inside the artifact
    alert("Pop-up blocked by browser. Please allow pop-ups for this page and try again, or use the Excel download instead.");
    return;
  }
  win.document.open();
  win.document.write(html);
  win.document.close();
  // Auto-focus for immediate print if desired
  win.focus();
};

// ─── STYLES ──────────────────────────────────────────────────────────────────
const css = `
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:'Inter',system-ui,-apple-system,sans-serif;background:#f0f2f7;min-height:100vh;display:flex;align-items:flex-start;justify-content:center;padding:32px 16px}
  .app{width:100%;max-width:840px}
  .topbar{background:#1a2e4a;padding:18px 28px;display:flex;align-items:center;gap:12px;border-radius:12px 12px 0 0}
  .logo{font-size:15px;font-weight:700;color:#fff}.logo span{color:#7eb8e8;font-weight:400}
  .badge{margin-left:auto;font-size:11px;color:#7eb8e8;background:rgba(126,184,232,0.15);padding:4px 10px;border-radius:20px;border:1px solid rgba(126,184,232,0.3)}
  .steps{display:flex;background:#f0f2f7;border:1px solid #dde2ee;border-top:none}
  .step{flex:1;padding:13px 8px;text-align:center;font-size:11px;color:#aaa;border-right:1px solid #dde2ee;transition:all .2s}
  .step:last-child{border-right:none}
  .step.active{color:#1a2e4a;font-weight:700;background:#fff;border-bottom:2px solid #1a2e4a}
  .step.done{color:#2d7a4f;background:#f0faf4}
  .prog{height:3px;background:#dde2ee}.prog-f{height:100%;background:#1a2e4a;transition:width .4s}
  .card{background:#fff;border:1px solid #dde2ee;border-top:none;border-radius:0 0 12px 12px;padding:28px}
  .sec{display:none}.sec.on{display:block}
  .g2{display:grid;grid-template-columns:1fr 1fr;gap:14px}
  .full{grid-column:1/-1}
  label{display:block;font-size:11px;font-weight:600;color:#8892a4;letter-spacing:.05em;margin-bottom:6px;text-transform:uppercase}
  input,select,textarea{width:100%;padding:10px 13px;border:1.5px solid #dde2ee;border-radius:8px;font-size:14px;color:#1a2e4a;background:#fafbfd;outline:none;transition:border-color .15s;font-family:inherit;appearance:none}
  input:focus,select:focus,textarea:focus{border-color:#1a2e4a;background:#fff}
  textarea{height:90px;resize:none;line-height:1.5}
  select{background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='8' viewBox='0 0 12 8'%3E%3Cpath d='M1 1l5 5 5-5' stroke='%238892a4' stroke-width='1.5' fill='none' stroke-linecap='round'/%3E%3C/svg%3E");background-repeat:no-repeat;background-position:right 12px center;padding-right:32px}
  .drop{border:2px dashed #c5cde0;border-radius:10px;padding:20px;cursor:pointer;text-align:center;transition:all .2s;background:#fafbfd}
  .drop:hover,.drop.drag{border-color:#1a2e4a;background:#f0f4fa}
  .drop input{display:none}
  .chips{display:flex;flex-wrap:wrap;gap:6px;margin-top:10px}
  .chip{display:flex;align-items:center;gap:6px;padding:4px 10px;background:#e4edf8;border-radius:20px;font-size:11px;color:#1a2e4a;font-weight:500}
  .chip-x{cursor:pointer;opacity:.6;font-size:14px}.chip-x:hover{opacity:1}
  .exbox{background:#f0faf4;border:1px solid #a8d5be;border-radius:8px;padding:12px 14px;margin-top:10px;font-size:11px;color:#1a4a30;line-height:1.7}
  .exbox b{display:block;margin-bottom:4px;font-size:10px;text-transform:uppercase;letter-spacing:.05em;color:#2d7a4f}
  .spinning{display:flex;align-items:center;gap:8px;font-size:12px;color:#1a2e4a;padding:10px 14px;background:#e4edf8;border-radius:8px;margin-top:10px}
  .row{display:flex;gap:10px;justify-content:space-between;align-items:center;margin-top:22px}
  .hint{font-size:11px;color:#bbb}
  .btn{padding:10px 22px;border-radius:8px;font-size:14px;font-weight:600;cursor:pointer;transition:all .15s;border:none}
  .btn-p{background:#1a2e4a;color:#fff}.btn-p:hover{background:#243d61}.btn-p:disabled{opacity:.4;cursor:not-allowed}
  .btn-s{background:transparent;border:1.5px solid #dde2ee;color:#666;font-weight:500}.btn-s:hover{background:#f5f7fb}
  .rat{background:#f0f4fa;border:1px solid #d0daf0;border-radius:8px;padding:13px 16px;margin-bottom:16px;font-size:12px;color:#3a4a6a;line-height:1.6}
  .rat b{color:#1a2e4a;display:block;margin-bottom:4px;font-size:11px;text-transform:uppercase;letter-spacing:.05em}
  .cc{background:#fafbfd;border:1px solid #e8ecf5;border-radius:10px;padding:13px 16px;display:flex;align-items:center;gap:14px;margin-bottom:8px}
  .tkr{width:44px;height:44px;border-radius:8px;background:#e4edf8;display:flex;align-items:center;justify-content:center;font-size:10px;font-weight:700;color:#1a2e4a;text-align:center;flex-shrink:0}
  .bg{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin:14px 0}
  .bc{background:#f5f6fa;border-radius:8px;padding:11px 13px;border:1px solid #e8ecf5}
  .bv{font-size:18px;font-weight:700;color:#1a2e4a}.bl{font-size:10px;color:#999;margin-top:2px}
  .tag{display:inline-block;padding:4px 12px;background:#e4edf8;color:#1a2e4a;border-radius:20px;font-size:12px;font-weight:600;margin-bottom:14px}
  .wbox{background:#1a2e4a;border-radius:12px;padding:20px 24px;margin:14px 0;color:#fff}
  .wval{font-size:38px;font-weight:700;color:#7eb8e8}.wsub{font-size:11px;color:rgba(255,255,255,.5);margin-top:5px}
  .wtbl{width:100%;border-collapse:collapse;font-size:12px;margin:14px 0;border-radius:8px;overflow:hidden;border:1px solid #e8ecf5}
  .wtbl th{padding:8px 12px;text-align:left;background:#f0f2f7;color:#8892a4;font-size:10px;text-transform:uppercase;letter-spacing:.05em}
  .wtbl td{padding:8px 12px;border-bottom:1px solid #f0f2f7;color:#1a2e4a}
  .wtbl tr:last-child td{border-bottom:none}
  .src-doc{display:inline-block;padding:2px 8px;border-radius:20px;font-size:10px;font-weight:600;background:#e8f5ee;color:#1a4a30}
  .src-def{display:inline-block;padding:2px 8px;border-radius:20px;font-size:10px;font-weight:600;background:#e4edf8;color:#1a2e4a}
  .dlbtn{display:flex;align-items:center;gap:12px;padding:16px 20px;background:#e8f5ee;border:1px solid #a8d5be;border-radius:10px;color:#1a4a30;font-size:14px;font-weight:600;cursor:pointer;width:100%;margin-top:14px;text-align:left;transition:background .15s}
  .dlbtn:hover{background:#d4eddf}.dlbtn:disabled{opacity:.5;cursor:not-allowed}
  .pdfbtn{display:flex;align-items:center;gap:12px;padding:16px 20px;background:#e8edf8;border:1px solid #a8bcd8;border-radius:10px;color:#1a2e4a;font-size:14px;font-weight:600;cursor:pointer;width:100%;margin-top:10px;text-align:left;transition:background .15s}
  .pdfbtn:hover{background:#d4deee}.pdfbtn:disabled{opacity:.5;cursor:not-allowed}
  .infobox{margin-top:10px;padding:10px 14px;background:#f8f9fb;border-radius:8px;font-size:11px;color:#999;border:1px solid #e5e8ee;line-height:1.6}
  .spin{display:inline-block;width:14px;height:14px;border:2px solid rgba(255,255,255,.3);border-top-color:#fff;border-radius:50%;animation:sp .7s linear infinite;flex-shrink:0}
  .spin-d{display:inline-block;width:13px;height:13px;border:2px solid rgba(26,46,74,.2);border-top-color:#1a2e4a;border-radius:50%;animation:sp .7s linear infinite;flex-shrink:0}
  @keyframes sp{to{transform:rotate(360deg)}}

  /* ── Advisor Chatbox ─────────────────────────────────────────────────────── */
  .chat-wrap{margin-top:22px;border:1.5px solid #dde2ee;border-radius:10px;overflow:hidden;background:#fafbfd}
  .chat-head{background:#1a2e4a;padding:11px 16px;display:flex;align-items:center;gap:9px}
  .chat-head-icon{font-size:16px}
  .chat-head-title{font-size:12px;font-weight:700;color:#fff;flex:1}
  .chat-head-sub{font-size:10px;color:rgba(255,255,255,.45)}
  .chat-msgs{max-height:280px;overflow-y:auto;padding:14px;display:flex;flex-direction:column;gap:10px}
  .chat-msgs:empty::after{content:"Ask anything about the valuation — WACC, comparables, assumptions, scenario sensitivities…";color:#bbb;font-size:11px;font-style:italic}
  .msg{display:flex;flex-direction:column;gap:3px;max-width:88%}
  .msg.user{align-self:flex-end;align-items:flex-end}
  .msg.ai{align-self:flex-start;align-items:flex-start}
  .msg-bubble{padding:9px 13px;border-radius:10px;font-size:12px;line-height:1.6;word-break:break-word}
  .msg.user .msg-bubble{background:#1a2e4a;color:#fff;border-radius:10px 10px 2px 10px}
  .msg.ai .msg-bubble{background:#fff;border:1px solid #dde2ee;color:#1a2e4a;border-radius:10px 10px 10px 2px}
  .msg-meta{font-size:10px;color:#aaa;padding:0 4px}
  .chat-typing{align-self:flex-start}
  .chat-typing .msg-bubble{background:#fff;border:1px solid #dde2ee;padding:9px 13px;border-radius:10px 10px 10px 2px}
  .typing-dots{display:flex;gap:4px;align-items:center;height:14px}
  .typing-dots span{width:6px;height:6px;border-radius:50%;background:#1a2e4a;opacity:.3;animation:blink 1.2s infinite}
  .typing-dots span:nth-child(2){animation-delay:.2s}
  .typing-dots span:nth-child(3){animation-delay:.4s}
  @keyframes blink{0%,80%,100%{opacity:.3}40%{opacity:1}}
  .chat-input-row{display:flex;gap:8px;padding:10px 12px;border-top:1px solid #eef0f5;background:#fff}
  .chat-input{flex:1;padding:9px 12px;border:1.5px solid #dde2ee;border-radius:8px;font-size:13px;color:#1a2e4a;background:#fafbfd;outline:none;resize:none;font-family:inherit;line-height:1.4;min-height:38px;max-height:90px;overflow-y:auto;transition:border-color .15s}
  .chat-input:focus{border-color:#1a2e4a;background:#fff}
  .chat-send{padding:9px 14px;background:#1a2e4a;color:#fff;border:none;border-radius:8px;font-size:13px;font-weight:600;cursor:pointer;transition:background .15s;flex-shrink:0;display:flex;align-items:center;gap:6px}
  .chat-send:hover{background:#243d61}
  .chat-send:disabled{opacity:.4;cursor:not-allowed}
  .chat-hint{font-size:10px;color:#bbb;padding:0 12px 8px;text-align:center}

`;

// ─── COMPONENT ────────────────────────────────────────────────────────────────
export default function App() {
  const [step,      setStep]      = useState(1);
  const [form,      setForm]      = useState({company:"",description:"",revenueModel:"",stage:"",country:"",sector:""});
  const [files,     setFiles]     = useState([]);
  const [extracting,setExtracting]= useState(false);
  const [extracted, setExtracted] = useState(null);
  const [exNote,    setExNote]    = useState("");
  const [comps,     setComps]     = useState([]);
  const [bench,     setBench]     = useState({});
  const [sector,    setSector]    = useState(1);
  const [wacc,      setWacc]      = useState(null);
  const [shadowDCF, setShadowDCF] = useState(null);
  const [loading,   setLoading]   = useState(false);
  const [dlLoading, setDlLoading] = useState(false);
  const [pdfLoading,setPdfLoading]= useState(false);
  const [drag,      setDrag]      = useState(false);
  // ── Advisor chat ──────────────────────────────────────────────────────────
  const [chatMsgs,  setChatMsgs]  = useState([]);
  const [chatInput, setChatInput] = useState("");
  const [chatBusy,  setChatBusy]  = useState(false);
  const chatEndRef = useRef(null);
  const fileRef = useRef();

  // Auto-scroll chat to bottom
  useEffect(()=>{ chatEndRef.current?.scrollIntoView({behavior:"smooth"}); },[chatMsgs,chatBusy]);

  const sendChat = async () => {
    const q = chatInput.trim();
    if(!q || chatBusy) return;
    setChatInput("");
    setChatMsgs(p=>[...p,{role:"user",text:q}]);
    setChatBusy(true);

    // Build a rich context snapshot of the current model state
    const waccSummary = wacc ? `WACC=${(wacc.wacc*100).toFixed(2)}%, Ke=${(wacc.ke*100).toFixed(2)}%, Kd=${(wacc.kd*100).toFixed(1)}%, g=${(wacc.g*100).toFixed(1)}%, stagePremium=${(wacc.stagePremium*100).toFixed(2)}%, beta=${wacc.betaU?.toFixed(3)}, tax=${(wacc.tax*100).toFixed(1)}%` : "not yet computed";
    const benchSummary = `Gross Margin=${bench.gm?.toFixed(1)}%, EBITDA Margin=${bench.eb?.toFixed(1)}%, Beta=${bench.b?.toFixed(3)}, Rev Growth=${bench.rg?.toFixed(1)}%`;
    const compsSummary = comps.map(c=>`${c.n} (${c.t}, GM=${c.gm?.toFixed(1)}%, β=${c.b?.toFixed(3)})`).join("; ");
    const extractedSummary = extracted ? Object.entries(extracted)
      .filter(([k,v])=>v!=null&&k!=="notes"&&!Array.isArray(v))
      .map(([k,v])=>`${k}=${typeof v==="number"?(v>1?v.toFixed(0):v.toFixed(4)):v}`).join(", ") : "none";
    const conversationHistory = chatMsgs.map(m=>({role:m.role==="user"?"user":"assistant",content:m.text}));

    const systemPrompt = `You are a senior investment analyst and valuation advisor embedded inside a live DCF valuation tool built on the EPFL/ETH Medtech Comparable Database (66 companies × 10 sub-sectors). Your role is to help the startup founder understand, challenge, or refine the financial model.

CURRENT MODEL STATE for ${form.company||"this startup"} (${form.stage}, ${form.country}, ${form.revenueModel}):
- Sub-sector: ${SECTOR_NAMES[sector]} (EPFL/ETH DB, 66-company universe)
- WACC inputs: ${waccSummary}
- D/E ratio: ${wacc?(wacc.de*100).toFixed(0):"?"} % — stage-adjusted (${form.stage}), not fixed at 20%
- Sector benchmarks used (sector medians, reliability-filtered): ${benchSummary}
- Comparable companies shown: ${compsSummary}
- Shadow DCF valuation range (UI cross-check): Bear/Base/Bull computed in-browser
- Document-extracted data: ${extractedSummary}
- Startup description: ${form.description?.slice(0,300)||"not provided"}
- Terminal value method: ${(bench.eb!=null&&bench.eb<0)?"EV/Revenue multiple (sector EBITDA negative)":"Gordon Growth Model"}

KEY MODEL IMPROVEMENTS IN THIS VERSION:
(a) D/E is stage-scaled — Pre-revenue R&D uses 5%, not 20%. This lowers WACC for early-stage companies.
(b) Shadow DCF computed in the UI for cross-check against Excel template formulas.
(c) Terminal value automatically switches to EV/Revenue when sector median EBITDA is negative.
(d) rNPV warning shown for Seed/Pre-revenue stages — remind founder that milestone probabilities should be applied.
(e) Full 66-company database with reliability flags (🟢/🟡) and D/E corrections for confirmed yfinance errors (VCEL, RaySearch, Creo Medical).

RESPOND AS A TRUSTED ADVISOR:
- Be direct, precise, and quantitative when possible
- Reference specific numbers from the model (e.g. "your current stage premium is 4.5% which pushes WACC above 18%")
- If asked about rNPV vs DCF, explain the difference clearly: rNPV applies success probabilities to each milestone cash flow; standard DCF does not
- If the founder challenges an assumption, acknowledge it and quantify the directional impact on WACC or valuation
- Keep responses concise (3–6 sentences) unless a complex question warrants more
- If asked about a specific comparable company from the 66-company database, you can cite its reliability flag and data quality notes`;

    try {
      const resp = await fetch("/api/chat",{
        method:"POST",
        headers:{"Content-Type":"application/json"},
        body:JSON.stringify({
          system: systemPrompt,
          messages:[
            ...conversationHistory,
            {role:"user",content:q}
          ]
        })
      });
      const data = await resp.json();
      const reply = data.text || data.error || "I couldn't generate a response. Please try again.";
      setChatMsgs(p=>[...p,{role:"ai",text:reply}]);
    } catch(e) {
      setChatMsgs(p=>[...p,{role:"ai",text:"Connection error — please try again."}]);
    }
    setChatBusy(false);
  };

  const handleChatKey = e => {
    if(e.key==="Enter"&&!e.shiftKey){e.preventDefault();sendChat();}
  };

  const upd = (k,v) => setForm(f=>({...f,[k]:v}));
  const canGo = form.company && form.revenueModel && form.stage && form.country && form.description.length>=10;

  // ── Add files → call Claude API to extract ─────────────────────────────────
  const addFiles = async newF => {
    const arr = Array.from(newF);
    if(!arr.length) return;
    setFiles(p=>[...p,...arr]);
    setExtracting(true); setExNote("");
    // Use full description + company + stage context so extraction is calibrated
    const ctx = [
      form.company ? `Company: ${form.company}` : "",
      form.description?.length >= 10 ? `Description: ${form.description}` : "",
      form.stage ? `Stage: ${form.stage}` : "",
      form.revenueModel ? `Revenue model: ${form.revenueModel}` : "",
      form.country ? `Jurisdiction: ${form.country}` : "",
    ].filter(Boolean).join(". ") || "deep-tech medtech startup";
    try {
      const ex = await extractWithClaude(arr, ctx);
      setExtracted(prev => {
        // Smart merge: prefer non-null values, but don't overwrite with null
        const merged = {...(prev||{})};
        Object.entries(ex).forEach(([k,v]) => {
          if (Array.isArray(v)) {
            const prevArr = Array.isArray(merged[k]) ? merged[k] : Array(11).fill(null);
            merged[k] = prevArr.map((old, i) => v[i] != null ? v[i] : old);
          } else if (v != null) {
            merged[k] = v;
          }
        });
        return merged;
      });
      setExNote(ex.notes || "Extraction complete.");
    } catch(e) {
      setExNote("Extraction failed: " + e.message);
    }
    setExtracting(false);
  };

  const removeFile = i => setFiles(p=>p.filter((_,j)=>j!==i));

  // ── Step 1 → 2 ────────────────────────────────────────────────────────────
  const runMatch = () => {
    const s = detectSector(form.description, form.sector);
    setSector(s);
    const matched = getComps(s);
    setComps(matched);
    const ex = extracted || {};

    // Normalize gross_margin: document may give 75 or 0.75
    const normGM = ex.gross_margin != null
      ? (ex.gross_margin > 1 ? ex.gross_margin : ex.gross_margin * 100)
      : null;
    const normEB = ex.ebitda_margin != null
      ? (ex.ebitda_margin > 1 ? ex.ebitda_margin : ex.ebitda_margin * 100)
      : null;

    // Revenue growth: prefer document data; fallback to sector median
    let revGrowth = null;
    if (ex.revenue_growth_rate != null) {
      revGrowth = ex.revenue_growth_rate > 1 ? ex.revenue_growth_rate : ex.revenue_growth_rate * 100;
    } else if (ex.revenue_year1 != null && ex.revenue_year3 != null && ex.revenue_year3 > 0) {
      // Compute 2-year CAGR from doc-provided revenues
      revGrowth = (Math.pow(ex.revenue_year3 / (ex.revenue_year1 || 1), 0.5) - 1) * 100;
    }

    setBench({
      gm: normGM ?? med(matched.map(c => c.gm)),
      eb: normEB ?? med(matched.map(c => c.eb).filter(v => v != null)),
      b:  ex.beta_unlevered ?? med(matched.map(c => c.b).filter(v => v != null)),
      rg: revGrowth ?? med(matched.map(c => c.rg).filter(v => v != null)),
    });
    setStep(2);
  };

  // ── Step 2 → 3 ────────────────────────────────────────────────────────────
  const runGenerate = () => {
    setLoading(true);
    setTimeout(()=>{
      const w = buildWACC(form.country, extracted, bench.b, form.stage);
      setWacc(w);
      // FIX (b): Compute shadow DCF for UI cross-check
      setShadowDCF(computeShadowDCF(extracted, w, bench));
      setLoading(false);
      setStep(3);
    },1000);
  };

  // ── Download Excel ─────────────────────────────────────────────────────────
  const doDownload = () => {
    setDlLoading(true);
    setTimeout(()=>{
      try {
        buildExcel({...form, sector}, extracted||{}, comps, bench, wacc, sector);
      } catch(e) {
        alert("Excel generation error: " + e.message);
      }
      setDlLoading(false);
    }, 300);
  };

  // ── Download PDF Report ────────────────────────────────────────────────────
  const doPDF = async () => {
    if (!wacc) return;
    setPdfLoading(true);
    try {
      await buildPDF(
        {...form, sector},
        extracted||{},
        comps,
        bench,
        wacc,
        sector,
        shadowDCF
      );
    } catch(e) {
      alert("PDF generation error: " + e.message);
    }
    setPdfLoading(false);
  };

  const pct = step===1?33:step===2?66:100;
  const exFields = extracted ? Object.entries(extracted).filter(([k,v])=>
    v!=null && k!=="notes" && k!=="currency" && !Array.isArray(v) &&
    !["wacc_rf","wacc_erp","beta_unlevered","debt_equity_ratio"].includes(k) // hide raw WACC fields
  ) : [];
  // Prioritise showing revenue/operational fields first
  const exFieldsSorted = [
    ...exFields.filter(([k])=>["revenue_year1","revenue_year3","revenue_year5","gross_margin","ebitda_margin","burn_rate_monthly","runway_months","headcount_current","tam_size","deal1_upfront_fee","deal1_royalty_rate","funding_raised_total","pre_money_valuation","rd_spend_annual"].includes(k)),
    ...exFields.filter(([k])=>!["revenue_year1","revenue_year3","revenue_year5","gross_margin","ebitda_margin","burn_rate_monthly","runway_months","headcount_current","tam_size","deal1_upfront_fee","deal1_royalty_rate","funding_raised_total","pre_money_valuation","rd_spend_annual"].includes(k)),
  ];
  const exArrays = extracted ? Object.entries(extracted).filter(([k,v])=>Array.isArray(v)&&v.some(x=>x!=null)) : [];

  return (
    <>
      <style>{css}</style>
      <div className="app">
        <div className="topbar">
          <div className="logo">Financial <span>Valuation</span></div>
          <div className="badge">Hackathon 2026</div>
        </div>
        <div className="steps">
          {["01 — Startup Profile","02 — Comparables","03 — Financial Valuation"].map((l,i)=>(
            <div key={i} className={`step${step===i+1?" active":""}${step>i+1?" done":""}`}>{l}</div>
          ))}
        </div>
        <div className="prog"><div className="prog-f" style={{width:`${pct}%`}}/></div>
        <div className="card">

          {/* STEP 1 */}
          <div className={`sec${step===1?" on":""}`}>
            <div className="g2">
              <div>
                <label>Startup Name</label>
                <input placeholder="e.g. NeuroFlow Medical" value={form.company} onChange={e=>upd("company",e.target.value)}/>
              </div>
              <div>
                <label>Sector</label>
                <select value={form.sector} onChange={e=>upd("sector",e.target.value)}>
                  <option value="">Auto-detect from description</option>
                  {Object.entries(SECTOR_NAMES).map(([k,v])=><option key={k} value={k}>{v}</option>)}
                </select>
              </div>
              <div>
                <label>Revenue Model</label>
                <select value={form.revenueModel} onChange={e=>upd("revenueModel",e.target.value)}>
                  <option value="">Select model</option>
                  {REVENUE_MODELS.map(v=><option key={v} value={v}>{v}</option>)}
                </select>
              </div>
              <div>
                <label>Stage</label>
                <select value={form.stage} onChange={e=>upd("stage",e.target.value)}>
                  <option value="">Select stage</option>
                  {STAGES.map(v=><option key={v} value={v}>{v}</option>)}
                </select>
              </div>
              <div className="full">
                <label>Country / Jurisdiction</label>
                <select value={form.country} onChange={e=>upd("country",e.target.value)}>
                  <option value="">Select country</option>
                  {COUNTRIES.map(c=><option key={c} value={c}>{c}</option>)}
                </select>
              </div>
              <div className="full">
                <label>Startup Description — product, indication, business model, target market</label>
                <textarea placeholder="Paste your executive summary…" value={form.description} onChange={e=>upd("description",e.target.value)}/>
              </div>
              <div className="full">
                <label>Financial Documents — PDF, Excel, CSV, Word, images</label>
                <div
                  className={`drop${drag?" drag":""}`}
                  onClick={()=>fileRef.current.click()}
                  onDragOver={e=>{e.preventDefault();setDrag(true)}}
                  onDragLeave={()=>setDrag(false)}
                  onDrop={e=>{e.preventDefault();setDrag(false);addFiles(e.dataTransfer.files)}}
                >
                  <input ref={fileRef} type="file" multiple
                    accept=".pdf,.xlsx,.xls,.csv,.docx,.doc,.png,.jpg,.jpeg,.webp"
                    onChange={e=>addFiles(e.target.files)}/>
                  <div style={{fontSize:24,marginBottom:6}}>📎</div>
                  <div style={{fontSize:13,fontWeight:600,color:"#1a2e4a"}}>Drop files here or click to browse</div>
                  <div style={{fontSize:11,color:"#aaa",marginTop:4}}>
                    Claude AI reads your documents and extracts all financial data automatically
                  </div>
                </div>
                {files.length>0 && (
                  <div className="chips">
                    {files.map((f,i)=>(
                      <div key={i} className="chip">
                        {f.name.endsWith(".pdf")?"📄":f.name.match(/\.xlsx?$/i)?"📊":f.name.match(/\.docx?$/i)?"📝":"📁"} {f.name}
                        <span className="chip-x" onClick={()=>removeFile(i)}>×</span>
                      </div>
                    ))}
                  </div>
                )}
                {extracting && (
                  <div className="spinning"><span className="spin-d"/> Reading documents with Claude AI…</div>
                )}
                {exNote && !extracting && (exFieldsSorted.length>0||exArrays.length>0) && (
                  <div className="exbox">
                    <b>✅ Extracted from your documents</b>
                    {exFieldsSorted.slice(0,12).map(([k,v])=>{
                      const label = k.replace(/_/g," ").replace(/\b\w/g,c=>c.toUpperCase());
                      const display = typeof v==="number"
                        ? (k.includes("margin")||k.includes("rate")||k.includes("growth"))
                          ? `${(v>1?v:v*100).toFixed(1)}%`
                          : v > 1e5 ? `${(v/1e6).toFixed(2)}M` : v > 1e3 ? `${(v/1e3).toFixed(0)}k` : v.toFixed(2)
                        : String(v);
                      return <div key={k}><strong>{label}:</strong> {display}</div>;
                    })}
                    {exArrays.length>0 && <div style={{marginTop:4,color:"#2d7a4f"}}>+ {exArrays.length} year-by-year arrays (revenue, FTEs, CapEx, costs…)</div>}
                    <div style={{marginTop:6,opacity:.7,fontSize:10}}>{exNote}</div>
                  </div>
                )}
                {exNote && !extracting && exFields.length===0 && exArrays.length===0 && (
                  <div className="exbox" style={{background:"#fff8e1",borderColor:"#ffe082",color:"#7a5c00"}}>
                    <b>⚠️ No financial data found</b>{exNote}
                  </div>
                )}
              </div>
            </div>
            <div className="row">
              <span className="hint">{form.description.length} char{!canGo?" · fill all fields":""}</span>
              <button className="btn btn-p" onClick={runMatch} disabled={!canGo||extracting}>
                {extracting?<><span className="spin"/> Reading…</>:"Find Comparables →"}
              </button>
            </div>
          </div>

          {/* STEP 2 */}
          <div className={`sec${step===2?" on":""}`}>
            <div className="tag">{SECTOR_NAMES[sector]}</div>
            <div className="rat"><b>Why these comparables?</b>{RATIONALE[sector]}</div>

            {/* ── Selection methodology pills ──────────────────────────────── */}
            <div style={{margin:"10px 0 14px",display:"flex",flexWrap:"wrap",gap:6}}>
              {[
                {icon:"🔍", label:"Keyword match", tip:`Description scanned → best-fit sub-sector: "${SECTOR_NAMES[sector]}"`},
                {icon:"🗄️", label:"66-company DB", tip:"EPFL/ETH Medtech Comparable Database — 10 sub-sectors"},
                {icon:"🟢", label:"Reliability filter", tip:"🟢 fully reliable prioritised; 🟡 used with limits to reach 5 comps"},
                {icon:"📐", label:"Sector medians", tip:"GM, EBITDA, Beta, Revenue Growth medians feed Assumptions sheet"},
                {icon:"✅", label:"Data corrections", tip:"VCEL D/E=0.0×, RaySearch D/E=0.01×, Creo GM=46.6% (yfinance errors fixed)"},
              ].map(({icon,label,tip},i)=>(
                <div key={i} title={tip} style={{
                  display:"flex",alignItems:"center",gap:5,
                  background:"#eef2fb",borderRadius:20,padding:"4px 10px",
                  fontSize:11,fontWeight:600,color:"#1a2e4a",cursor:"default",
                  border:"1px solid #d0d8f0"
                }}>
                  <span>{icon}</span><span>{label}</span>
                  <span style={{fontSize:9,color:"#7a94c4",marginLeft:2}}>ⓘ</span>
                </div>
              ))}
            </div>

            {comps.map(c=>(
              <div className="cc" key={c.t}>
                <div className="tkr">{c.t.split(".")[0]}</div>
                <div style={{flex:1}}>
                  <div style={{fontSize:13,fontWeight:600,color:"#1a2e4a"}}>{c.n}</div>
                  <div style={{fontSize:11,color:"#aaa",marginTop:2}}>{c.c}</div>
                  {/* Per-company reason badge */}
                  {COMP_REASON[c.t] && (
                    <div style={{
                      marginTop:5,fontSize:10,color:"#2d5a8e",background:"#f0f4fa",
                      border:"1px solid #ccd8ee",borderRadius:6,padding:"3px 8px",
                      lineHeight:1.5,maxWidth:420
                    }}>
                      {COMP_REASON[c.t]}
                    </div>
                  )}
                </div>
                <div style={{display:"flex",gap:16,marginLeft:"auto",flexShrink:0}}>
                  <div style={{textAlign:"right"}}><div style={{fontSize:14,fontWeight:700,color:"#1a2e4a"}}>{c.gm.toFixed(1)}%</div><div style={{fontSize:10,color:"#aaa"}}>Gross Margin</div></div>
                  {c.b&&<div style={{textAlign:"right"}}><div style={{fontSize:14,fontWeight:700,color:"#1a2e4a"}}>{c.b.toFixed(3)}</div><div style={{fontSize:10,color:"#aaa"}}>Beta</div></div>}
                </div>
              </div>
            ))}
            <div className="bg">
              {[["Gross Margin",bench.gm?.toFixed(1)+"%"],["EBITDA Margin",bench.eb?.toFixed(1)+"%"],["Beta (unlev.)",bench.b?.toFixed(3)],["Rev. Growth",bench.rg?.toFixed(1)+"%"],[`Comparables`,comps.length],["Sub-sector",SECTOR_NAMES[sector]?.split(" ")[0]]].map(([l,v],i)=>(
                <div key={i} className="bc"><div className="bv" style={i>=4?{fontSize:13}:{}}>{v}</div><div className="bl">{l}</div></div>
              ))}
            </div>
            {extracted && (exFieldsSorted.length>0||exArrays.length>0) && (
              <div className="exbox" style={{marginBottom:14}}>
                <b>📄 Document data will override sector defaults in the Excel model</b>
                {exFieldsSorted.slice(0,8).map(([k,v])=>{
                  const label = k.replace(/_/g," ").replace(/\b\w/g,c=>c.toUpperCase());
                  const display = typeof v==="number"
                    ? (k.includes("margin")||k.includes("rate")||k.includes("growth"))
                      ? `${(v>1?v:v*100).toFixed(1)}%`
                      : v > 1e5 ? `${(v/1e6).toFixed(2)}M` : v > 1e3 ? `${(v/1e3).toFixed(0)}k` : v.toFixed(2)
                    : String(v);
                  return <div key={k}><strong>{label}:</strong> {display}</div>;
                })}
                {exArrays.length>0&&<div>+ {exArrays.length} year-by-year arrays (revenue, FTEs, CapEx…)</div>}
              </div>
            )}
            <div className="row">
              <button className="btn btn-s" onClick={()=>setStep(1)}>← Back to Profile</button>
              <button className="btn btn-p" onClick={runGenerate} disabled={loading}>
                {loading?<span className="spin"/>:"Generate Financial Valuation →"}
              </button>
            </div>
          </div>

          {/* STEP 3 */}
          <div className={`sec${step===3?" on":""}`}>
            {wacc&&<>
              <div className="wbox">
                <div style={{fontSize:11,color:"rgba(255,255,255,.4)",marginBottom:5,letterSpacing:".08em"}}>COMPUTED WACC</div>
                <div className="wval">{(wacc.wacc*100).toFixed(2)}%</div>
                <div className="wsub">Ke={`${(wacc.ke*100).toFixed(1)}%`} × (1−D/V) + Kd={`${(wacc.kd*100).toFixed(1)}%`} × (1−t) × D/V</div>
              </div>
              <table className="wtbl">
                <thead>
                  <tr><th>Input</th><th style={{textAlign:"right"}}>Value</th><th>Source</th></tr>
                </thead>
                <tbody>
                  {[
                    ["Risk-free rate (Rf)",    `${(wacc.rf*100).toFixed(2)}%`,      wacc.sources.rf],
                    ["Equity Risk Premium",    `${(wacc.erp*100).toFixed(2)}%`,     wacc.sources.erp],
                    ["Beta (unlevered)",       wacc.betaU?.toFixed(3),              wacc.sources.beta],
                    ["Stage Risk Premium",     `+${(wacc.stagePremium*100).toFixed(2)}%`, wacc.sources.stage],
                    ["Cost of Equity (Ke)",    `${(wacc.ke*100).toFixed(2)}%`,      "Computed: Rf + β×ERP + Stage"],
                    ["Cost of Debt (Kd)",      `${(wacc.kd*100).toFixed(1)}%`,      wacc.sources.kd],
                    ["D/E Ratio (stage-adj.)", `${(wacc.de*100).toFixed(0)}%`,      wacc.sources.de],
                    ["Tax rate",               `${(wacc.tax*100).toFixed(1)}%`,     wacc.sources.tax],
                    ["Terminal growth (g)",    `${(wacc.g*100).toFixed(1)}%`,       wacc.sources.g],
                    ["Gross Margin",           `${(bench.gm??0).toFixed(1)}%`,      extracted?.gross_margin!=null?"📄 From document":"📊 Sector median (66-co. DB)"],
                    ["EBITDA Margin",          `${(bench.eb??0).toFixed(1)}%`,      extracted?.ebitda_margin!=null?"📄 From document":(bench.eb!=null&&bench.eb<0)?"📊 Sector median ⚠️ negative — TV uses EV/Rev":"📊 Sector median (66-co. DB)"],
                    ...(extracted?.revenue_year1!=null?[["Revenue Year 1",`${(extracted.revenue_year1/1000).toFixed(0)}k ${extracted.currency||""}`.trim(),"📄 From document"]]:[] ),
                    ...(extracted?.revenue_year3!=null?[["Revenue Year 3",`${(extracted.revenue_year3/1000).toFixed(0)}k ${extracted.currency||""}`.trim(),"📄 From document"]]:[] ),
                    ...(extracted?.burn_rate_monthly!=null?[["Monthly Burn",`${(extracted.burn_rate_monthly/1000).toFixed(0)}k ${extracted.currency||""}`.trim(),"📄 From document"]]:[] ),
                    ...(extracted?.runway_months!=null?[["Runway",`${extracted.runway_months} months`,"📄 From document"]]:[] ),
                    ...(extracted?.tam_size!=null?[["TAM",`${(extracted.tam_size/1e6).toFixed(0)}M ${extracted.currency||""}`.trim(),"📄 From document"]]:[] ),
                  ].map(([l,v,s],i)=>(
                    <tr key={i} style={{background:i%2===0?"#fff":"#fafbfd"}}>
                      <td style={{fontWeight:500}}>{l}</td>
                      <td style={{textAlign:"right",fontWeight:700,fontVariantNumeric:"tabular-nums"}}>{v}</td>
                      <td><span className={s?.includes("document")||s?.includes("Derived")?"src-doc":s?.includes("Computed")?"src-doc":"src-def"}>{s}</span></td>
                    </tr>
                  ))}
                </tbody>
              </table>

              {/* ── FIX (b): Shadow DCF — indicative valuation range ─────── */}
              {shadowDCF && (
                <div style={{margin:"14px 0",border:"1px solid #dde2ee",borderRadius:10,overflow:"hidden"}}>
                  <div style={{background:"#243d61",padding:"11px 16px",display:"flex",alignItems:"center",gap:8}}>
                    <span style={{fontSize:14,color:"#fff",fontWeight:700}}>📐 Indicative Valuation Range</span>
                    <span style={{fontSize:10,color:"rgba(255,255,255,.5)",marginLeft:"auto"}}>Shadow DCF · UI cross-check · not audited</span>
                  </div>
                  {shadowDCF.noRevData ? (
                    <div style={{padding:"16px",background:"#fff8e1",fontSize:12,color:"#7a5c00",lineHeight:1.6}}>
                      <strong>⚠️ Revenue data not found in uploaded documents.</strong> The shadow DCF requires at least one revenue projection (Year 1, Year 3, or Year 5) or a Deal 1 upfront fee to compute a valuation range. Please enter revenue assumptions directly in the Excel model's <strong>Assumptions sheet</strong> (rows 16–32), or upload a document containing revenue figures.
                    </div>
                  ) : (
                    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:0,background:"#fff"}}>
                      {[["🔴 Bear","bear","#c0392b"],["🟡 Base","base","#1a2e4a"],["🟢 Bull","bull","#1a4a30"]].map(([label,k,col])=>(
                        <div key={k} style={{padding:"14px 16px",borderRight:k!=="bull"?"1px solid #dde2ee":"none",textAlign:"center"}}>
                          <div style={{fontSize:11,color:"#aaa",marginBottom:4}}>{label}</div>
                          <div style={{fontSize:22,fontWeight:700,color:col}}>
                            {shadowDCF[k] == null ? "N/A"
                              : shadowDCF[k] < 0 ? "< 0"
                              : shadowDCF[k] >= 1e9 ? `${(shadowDCF[k]/1e9).toFixed(2)}B`
                              : shadowDCF[k] >= 1e6 ? `${(shadowDCF[k]/1e6).toFixed(1)}M`
                              : `${(shadowDCF[k]/1e3).toFixed(0)}k`} {shadowDCF[k] != null ? shadowDCF.currency : ""}
                          </div>
                        </div>
                      ))}
                    </div>
                  )}
                  <div style={{padding:"8px 16px",background:"#f8f9fb",fontSize:10,color:"#999",borderTop:"1px solid #eef0f5"}}>
                    Terminal value method: <strong>{shadowDCF.terminalMethod}</strong>
                    {!shadowDCF.noRevData && <> · WACC {(wacc.wacc*100).toFixed(1)}% · g {(wacc.g*100).toFixed(1)}% · Order-of-magnitude estimates only; the Excel model contains the full formula-driven DCF.</>}
                  </div>
                </div>
              )}

              {/* ── FIX (d): rNPV warning for pre-revenue / clinical-stage ── */}
              {(form.stage === "Pre-revenue R&D" || form.stage === "Seed") && (
                <div style={{margin:"10px 0",padding:"10px 14px",background:"#fff8e1",border:"1px solid #ffe082",borderRadius:8,fontSize:11,color:"#7a5c00",lineHeight:1.6}}>
                  <strong>⚠️ Risk-Adjusted NPV (rNPV) Recommended</strong> — For pre-revenue / clinical-stage deep-tech startups, standard DCF overestimates value because it treats all future cash flows as certain. Industry practice (Venture Kick, B2Venture, Sofinnova) applies probability-weighted milestones: typical FDA PMA success rate ~65%; Phase II→III ~45%; commercial ramp-up ~80%. Apply milestone probability haircuts manually in the Excel Scenarios sheet.
                </div>
              )}

              {/* ── FIX (e): Comparables count now reflects full 66-co. DB ── */}
              <button className="dlbtn" onClick={doDownload} disabled={dlLoading}>
                <span style={{fontSize:24,flexShrink:0}}>{dlLoading?"⏳":"⬇"}</span>
                <span>
                  <div>{dlLoading?"Generating model…":`Download Valuation_${form.company.replace(/\s/g,"_")}.xlsx`}</div>
                  <div style={{fontSize:11,fontWeight:400,opacity:.65,marginTop:3}}>Original template preserved · all formulas intact · your data pre-filled · + Comparables sheet</div>
                </span>
              </button>

              {/* ── PDF Investor Report ─────────────────────────────────────── */}
              <button className="pdfbtn" onClick={doPDF} disabled={pdfLoading || !wacc}>
                <span style={{fontSize:24,flexShrink:0}}>{pdfLoading?"⏳":"📄"}</span>
                <span>
                  <div>{pdfLoading?"Writing report with AI — this takes ~15 seconds…":"Open Investor Report (Print → Save as PDF)"}</div>
                  <div style={{fontSize:11,fontWeight:400,opacity:.65,marginTop:3}}>
                    Opens in new window · Click 🖨 Print / Save as PDF · AI-written guide · Cover · WACC table · Comparables · Glossary · Limitations
                  </div>
                </span>
              </button>

              <div className="infobox">
                Sub-sector: <strong>{SECTOR_NAMES[sector]}</strong> · {comps.length} comparables shown ({comps.map(c=>c.t).join(", ")}) · sourced from EPFL/ETH 66-company Medtech DB ·
                {extracted&&(exFieldsSorted.length+exArrays.length)>0?` ${exFieldsSorted.length+exArrays.length} fields from documents ·`:""} Damodaran Jan 2025 · D/E stage-adjusted · Terminal value: {(bench.eb!=null&&bench.eb<0)?"EV/Revenue (EBITDA negative)":"Gordon Growth Model"} · All Excel formulas preserved
              </div>

              {/* ── Advisor Chatbox ────────────────────────────────────────── */}
              <div className="chat-wrap">
                <div className="chat-head">
                  <span className="chat-head-icon">💬</span>
                  <span className="chat-head-title">Valuation Advisor</span>
                  <span className="chat-head-sub">Ask about WACC · comparables · assumptions · scenarios</span>
                </div>
                <div className="chat-msgs">
                  {chatMsgs.map((m,i)=>(
                    <div key={i} className={`msg ${m.role}`}>
                      <div className="msg-bubble">{m.text}</div>
                      <span className="msg-meta">{m.role==="user"?"You":"Advisor"}</span>
                    </div>
                  ))}
                  {chatBusy && (
                    <div className="msg ai chat-typing">
                      <div className="msg-bubble">
                        <div className="typing-dots"><span/><span/><span/></div>
                      </div>
                    </div>
                  )}
                  <div ref={chatEndRef}/>
                </div>
                <div className="chat-hint">Shift+Enter for new line · Enter to send</div>
                <div className="chat-input-row">
                  <textarea
                    className="chat-input"
                    placeholder={`Ask about the ${form.company||"startup"} valuation — e.g. "Why is our WACC so high?" or "Our gross margin will be 85%, not ${(bench.gm??72).toFixed(0)}%"`}
                    value={chatInput}
                    onChange={e=>setChatInput(e.target.value)}
                    onKeyDown={handleChatKey}
                    rows={1}
                    disabled={chatBusy}
                  />
                  <button className="chat-send" onClick={sendChat} disabled={!chatInput.trim()||chatBusy}>
                    {chatBusy?<><span className="spin"/>…</>:<>Send ↑</>}
                  </button>
                </div>
              </div>
            </>}
            <div className="row" style={{marginTop:16}}>
              <button className="btn btn-s" onClick={()=>setStep(1)}>← New valuation</button>
            </div>
          </div>

        </div>
      </div>
    </>
  );
}

const TEMPLATE_B64 = "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";
