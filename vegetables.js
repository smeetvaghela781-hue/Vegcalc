// Comprehensive Vegetable Database with Gujarati, Hindi, and English names
const VEGETABLE_CATALOG = [
  { id: "potato", en: "Potato", gu: "Batata (બટાટા)", hi: "Aloo (आलू)", defaultRetailMargin: 0.35, icon: "🥔" },
  { id: "onion", en: "Onion", gu: "Dungri (ડુંગળી)", hi: "Pyaz (प्याज)", defaultRetailMargin: 0.30, icon: "🧅" },
  { id: "tomato", en: "Tomato (Tameta)", gu: "Tameta (ટામેટા)", hi: "Tamatar (टमाटर)", defaultRetailMargin: 0.40, icon: "🍅" },
  { id: "chilli", en: "Green Chilli (Marcha)", gu: "Marcha (મરચાં)", hi: "Hari Mirch (हरी मिर्च)", defaultRetailMargin: 0.35, icon: "🌶️" },
  { id: "bottle_gourd", en: "Bottle Gourd (Dudhi)", gu: "Dudhi (દૂધી)", hi: "Lauki (लौकी)", defaultRetailMargin: 0.40, icon: "🥒" },
  { id: "cauliflower", en: "Cauliflower (Fulavar)", gu: "Fulavar (ફુલાવર)", hi: "Phool Gobhi (फूलगोभी)", defaultRetailMargin: 0.35, icon: "🥦" },
  { id: "okra", en: "Ladyfinger (Bhindo)", gu: "Bhindo (ભીંડો)", hi: "Bhindi (भिंडी)", defaultRetailMargin: 0.35, icon: "🌱" },
  { id: "brinjal", en: "Brinjal (Ringna)", gu: "Ringna (રીંગણા)", hi: "Baingan (बैंगन)", defaultRetailMargin: 0.40, icon: "🍆" },
  { id: "coriander", en: "Coriander (Kothmari)", gu: "Kothmari (કોથમીર)", hi: "Hara Dhaniya (धनिया)", defaultRetailMargin: 0.45, icon: "🌿" },
  { id: "lemon", en: "Lemon (Limbu)", gu: "Limbu (લીંબુ)", hi: "Nimbu (नींबू)", defaultRetailMargin: 0.35, icon: "🍋" },
  { id: "mint", en: "Mint (Phudino)", gu: "Phudino (ફુદીનો)", hi: "Pudina (पुदीना)", defaultRetailMargin: 0.40, icon: "🍃" },
  { id: "cluster_beans", en: "Cluster Beans (Guvar)", gu: "Guvar (ગુવાર)", hi: "Gawar (ग्वार)", defaultRetailMargin: 0.40, icon: "🌿" },
  { id: "cabbage", en: "Cabbage (Kobi)", gu: "Kobi (કોબી)", hi: "Patta Gobhi (पत्तागोभी)", defaultRetailMargin: 0.35, icon: "🥬" },
  { id: "carrot", en: "Carrot (Gajar)", gu: "Gajar (ગાજર)", hi: "Gajar (गाजर)", defaultRetailMargin: 0.35, icon: "🥕" },
  { id: "cucumber", en: "Cucumber (Kakdi)", gu: "Kakdi (કાકડી)", hi: "Kheera (खीरा)", defaultRetailMargin: 0.40, icon: "🥒" },
  { id: "ginger", en: "Ginger (Aadu)", gu: "Aadu (આદુ)", hi: "Adrak (अदरक)", defaultRetailMargin: 0.30, icon: "🫚" },
  { id: "garlic", en: "Garlic (Lassan)", gu: "Lassan (લસણ)", hi: "Lehsun (लहसुन)", defaultRetailMargin: 0.30, icon: "🧄" },
  { id: "spinach", en: "Spinach (Palak)", gu: "Palak (પાલક)", hi: "Palak (पालक)", defaultRetailMargin: 0.40, icon: "🥬" },
  { id: "fenugreek", en: "Fenugreek (Methi)", gu: "Methi (મેથી)", hi: "Methi (मेथी)", defaultRetailMargin: 0.40, icon: "🌱" },
  { id: "bitter_gourd", en: "Bitter Gourd (Karela)", gu: "Karela (કારેલા)", hi: "Karela (करेला)", defaultRetailMargin: 0.40, icon: "🥒" },
  { id: "tindora", en: "Ivy Gourd (Giloda)", gu: "Giloda / Tindora (ટીંડોળા)", hi: "Kundru (कुंदरू)", defaultRetailMargin: 0.35, icon: "🫛" },
  { id: "ridge_gourd", en: "Ridge Gourd (Turia)", gu: "Turia (તૂરિયાં)", hi: "Tori (तोरी)", defaultRetailMargin: 0.35, icon: "🥒" },
  { id: "sweet_potato", en: "Sweet Potato (Shakariya)", gu: "Shakariya (શક્કરિયાં)", hi: "Shakarkand (शकरकंद)", defaultRetailMargin: 0.35, icon: "🍠" },
  { id: "parwal", en: "Pointed Gourd (Parval)", gu: "Parval (પરવળ)", hi: "Parwal (परवल)", defaultRetailMargin: 0.35, icon: "🥒" },
  { id: "green_peas", en: "Green Peas (Vatana)", gu: "Vatana (વટાણા)", hi: "Matar (मटर)", defaultRetailMargin: 0.35, icon: "🫛" },
  { id: "capsicum", en: "Capsicum (Simla Marcha)", gu: "Simla Marcha (શિમલા મરચાં)", hi: "Shimla Mirch (शिमला मिर्च)", defaultRetailMargin: 0.35, icon: "🫑" },
  { id: "radish", en: "Radish (Mula)", gu: "Mula (મૂળા)", hi: "Mooli (मूली)", defaultRetailMargin: 0.40, icon: "🥢" },
  { id: "beetroot", en: "Beetroot (Beet)", gu: "Beet (બીટ)", hi: "Chukandar (चुकंदर)", defaultRetailMargin: 0.35, icon: "🟣" }
];

// Initial preloaded data from the user's Veg_Business_Tracker.xlsx Intake sheet
const DEFAULT_INITIAL_ENTRIES = [
  {
    id: "v-row-1",
    date: "2026-10-22",
    item: "Potato",
    qty: 50,
    totalPrice: 850,
    retailPrice: 25,
    supplier: "Ramesh Mandi"
  },
  {
    id: "v-row-2",
    date: "2026-10-22",
    item: "marcha",
    qty: 3,
    totalPrice: 150,
    retailPrice: 70,
    supplier: ""
  },
  {
    id: "v-row-3",
    date: "2026-10-22",
    item: "tameta",
    qty: 10,
    totalPrice: 75,
    retailPrice: 15,
    supplier: ""
  },
  {
    id: "v-row-4",
    date: "2026-10-22",
    item: "dudhi",
    qty: 5,
    totalPrice: 80,
    retailPrice: 25,
    supplier: ""
  },
  {
    id: "v-row-5",
    date: "2026-10-22",
    item: "fulavar",
    qty: 5,
    totalPrice: 100,
    retailPrice: 35,
    supplier: ""
  },
  {
    id: "v-row-6",
    date: "2026-10-22",
    item: "bhindo",
    qty: 3,
    totalPrice: 160,
    retailPrice: 70,
    supplier: ""
  },
  {
    id: "v-row-7",
    date: "2026-10-22",
    item: "Ringna",
    qty: 5,
    totalPrice: 80,
    retailPrice: 25,
    supplier: ""
  },
  {
    id: "v-row-8",
    date: "2026-10-22",
    item: "kothmari",
    qty: 2,
    totalPrice: 160,
    retailPrice: 120,
    supplier: ""
  },
  {
    id: "v-row-9",
    date: "2026-10-22",
    item: "limbu",
    qty: 2,
    totalPrice: 360,
    retailPrice: 240,
    supplier: ""
  },
  {
    id: "v-row-10",
    date: "2026-10-22",
    item: "phudino",
    qty: 1,
    totalPrice: 50,
    retailPrice: 80,
    supplier: ""
  },
  {
    id: "v-row-11",
    date: "2026-10-22",
    item: "Guvar",
    qty: 3,
    totalPrice: 120,
    retailPrice: 60,
    supplier: ""
  }
];
