'use client';

import React, { useState, useEffect, useMemo, useRef } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import * as XLSX from 'xlsx';
import { 
  Briefcase, 
  Upload, 
  QrCode, 
  PlusCircle, 
  CheckCircle, 
  AlertTriangle, 
  Clock, 
  Trash2, 
  Edit, 
  Search, 
  Filter, 
  RotateCcw, 
  Lock, 
  LogOut, 
  User, 
  AlertCircle, 
  Download,
  Check,
  FileSpreadsheet,
  Calendar,
  X,
  FileText,
  AlertOctagon,
  Menu,
  LayoutDashboard,
  Database,
  Sparkles,
  HelpCircle,
  RefreshCw,
  Settings,
  Camera,
  Keyboard,
  ClipboardList,
  ArrowRightCircle,
  Eye,
  Copy,
  ArrowLeft,
  ArrowUp,
  ArrowDown,
  Share2,
  Mail,
  Smartphone
} from 'lucide-react';

import { DEFAULT_IATA_AIRLINE_MAP, getCanonicalTag, get10DigitTag, matchTag } from './lib/iata';
import ScannerModal from './components/ScannerModal';

// Configurable Flight Numbers
const ARRIVAL_FLIGHTS = ['LH760', 'LH762', 'LX146', 'LX2646'];
const DEPARTURE_FLIGHTS = ['LH763', 'LH761', 'LX147', 'LX2647'];

// Interfaces based on BDO Format
interface BaggageRecord {
  id: string;
  sno: string;
  pir: string;
  name: string;
  originalTag: string;
  rushTag: string;
  flightNo: string; // Flight LH760, LH762, LX146, and LX2646, or others
  seal: string;
  ln: string; // Will store "Yes" or "No" (representing Locked)
  destination: string;
  remarks: string;
  storageRemarks?: string; // Remarks in Storage Location
  
  // Registry
  registryType: 'Arrival' | 'Departure';
  
  // Departure Specific
  departureForwardingStatus?: 'Awaiting Forwarding' | 'Forwarding Sent';
  departureStorageLocation?: 'LHG Office' | 'BMA' | 'Level 4';
  leftBehindReason?: string;
  forwardingFlightNo?: string;
  forwardingDestination?: string;
  forwardingDate?: string;
  forwardedBy?: string;
  forwardingRemarks?: string;
  specificStorageLocation?: string;
  bmaStorageLocation?: string;

  // Operational States
  status: 'Expected' | 'Received';
  receivedAt?: string; // ISO String (can be backdated for testing alerts)
  
  // Customs State (Arrival Only)
  customsStatus: 'Pending' | 'Cleared' | 'Not Cleared' | 'Marked Preventive';
  customsReason?: 'Lack of documents' | 'Awaiting documents' | 'Refused' | 'Deferred' | 'Preventive' | '';
  customsUpdatedAt?: string;
  
  // Disposition / Location State
  disposition: 'Pending' | 'Storage' | 'Delivered' | 'Forwarded' | 'Belt 9' | 'Handover' | 'CWC' | 'Re-export' | 'Awaiting Forwarding' | 'Forwarded on LHG Flight' | 'Collected by Passenger' | 'Returned to Airline' | 'Other';
  dispositionLocation?: 'Belt 9' | 'LHG Office' | 'BMA' | 'Level 4 Checks' | 'CWC' | 'VVM' | 'Outlook' | 'Advik' | 'Air India' | 'Indigo' | 'Spice Jet' | 'Hub Re-export' | 'Other Airline' | 'Level 4' | '';
  dispositionUpdatedAt?: string;
  
  createdAt: string;

  // New Fields for Enhanced Dispositions & Customs Operations Protocol
  weight?: number;
  damaged?: 'Y' | 'N';
  protocol?: 'Cleared Baggage' | 'Non-Cleared / Other' | '';
  deliveryAgent?: 'VVM' | 'Outlook' | 'Advik' | '';
  storageOption?: 'Standard Warehousing – LHG Office' | '';
  domesticForwarding?: 'Air India' | 'IndiGo' | 'SpiceJet' | 'No Forwarding' | '';
  arrivalBelt?: 'Arrival Belt 9' | '';
  handoverOption?: 'Partner Airlines' | '';
  warehouseOption?: 'CWC Warehouse' | '';
  reexportOption?: 'Re-export to Carrier Hub' | '';

  // Redesign fields
  preventiveReason?: string;
  handoverAirline?: string;
  handoverDateTime?: string;
  warehouseReference?: string;
  reexportFlight?: string;
  reexportDestination?: string;
  reexportDate?: string;
  deliveryDateTime?: string;
}

interface DictionaryEntry {
  field: string;
  label: string;
  aliases: string[];
  isMandatory: boolean;
  description: string;
}

const DEFAULT_MAPPING_DICTIONARY: DictionaryEntry[] = [
  { field: 'flightNo', label: 'Flight No', aliases: ['flight no', 'flight number', 'flight#', 'flight', 'flt', 'flt no', 'fltno', 'carrier', 'outbound flight number'], isMandatory: true, description: 'Carrier flight code (e.g. LH760)' },
  { field: 'originalTag', label: 'Original Tag', aliases: ['original tag', 'original bag tag', 'orig tag', 'org tag', 'orignal tag', 'original tga', 'bag tag', 'tag no', 'tag', 'originaltag', 'baggage tag', 'bag tag number'], isMandatory: false, description: '10-digit primary tag number' },
  { field: 'rushTag', label: 'Rush Tag', aliases: ['rush tag', 'rushtag', 'rush bag tag', 'r tag', 'rtag', 'expedite tag', 'expedite'], isMandatory: false, description: 'Rush/expedite tag number (e.g. LX920394)' },
  { field: 'name', label: 'Passenger Name', aliases: ['name', 'passenger', 'passenger name', 'pax', 'pax name', 'customer name', 'passengername', 'full name', 'pax fullname'], isMandatory: true, description: 'Passenger full name' },
  { field: 'pir', label: 'PIR Number', aliases: ['pir', 'pir no', 'pir number', 'pir ref', 'reference', 'pirnumber', 'report number', 'irregularity report'], isMandatory: true, description: 'Property Irregularity Report (e.g. DELEK12345)' },
  { field: 'weight', label: 'Weight (kg)', aliases: ['weight', 'wt', 'kg', 'kgs', 'weight kg', 'bag weight'], isMandatory: false, description: 'Baggage weight in kilograms (leave blank if not mentioned)' },
  { field: 'damaged', label: 'Damaged', aliases: ['damage', 'damaged', 'dmg', 'damage yn', 'condition', 'is damaged'], isMandatory: true, description: 'Damaged indicator (Y/N)' },
  { field: 'ln', label: 'Locked', aliases: ['locked', 'lock', 'l/n', 'ln', 'lock status', 'locked status', 'l.n'], isMandatory: false, description: 'Locked indicator (free text such as PAD, CL, Y, N)' },
  { field: 'destination', label: 'Destination', aliases: ['dest', 'destination', 'airport', 'dest code', 'station'], isMandatory: false, description: 'Three-letter airport code (e.g. DEL)' },
  { field: 'seal', label: 'Seal', aliases: ['seal', 'seal number', 'seal#', 'seals'], isMandatory: false, description: 'Customs or airline security seal number' },
  { field: 'remarks', label: 'Remarks', aliases: ['remark', 'remarks', 'notes', 'comment', 'comments', 'additional details'], isMandatory: false, description: 'Custom text remarks or instructions' },
  { field: 'protocol', label: 'Customs Protocol', aliases: ['protocol', 'customs protocol', 'disposition protocol', 'ops protocol', 'workflow protocol'], isMandatory: true, description: 'Cleared Baggage or Non-Cleared / Other' },
  { field: 'receivedAt', label: 'Date of Baggage Receipt', aliases: ['receipt date', 'registration date', 'date received', 'arrival date', 'received date', 'receipt'], isMandatory: false, description: 'Date and time the baggage was received' }
];

const DEFAULT_LOCK_DICTIONARY: Record<string, string> = {
  'CL': 'Combination Lock',
  'PAD': 'Padlock',
  'Y': 'Yes',
  'N': 'No',
  'SEAL': 'Sealed'
};

function getSimilarityScore(s1: string, s2: string): number {
  const clean = (s: string) => s.toLowerCase().replace(/[^a-z0-9]/g, '');
  const c1 = clean(s1);
  const c2 = clean(s2);
  if (c1 === c2) return 1.0;
  if (!c1 || !c2) return 0.0;
  if (c1.includes(c2) || c2.includes(c1)) {
    return Math.min(c1.length, c2.length) / Math.max(c1.length, c2.length) * 0.9;
  }
  
  // Levenshtein Distance
  const track = Array(c2.length + 1).fill(null).map(() => Array(c1.length + 1).fill(null));
  for (let i = 0; i <= c1.length; i += 1) track[0][i] = i;
  for (let j = 0; j <= c2.length; j += 1) track[j][0] = j;
  for (let j = 1; j <= c2.length; j += 1) {
    for (let i = 1; i <= c1.length; i += 1) {
      const indicator = c1[i - 1] === c2[j - 1] ? 0 : 1;
      track[j][i] = Math.min(
        track[j][i - 1] + 1, // deletion
        track[j - 1][i] + 1, // insertion
        track[j - 1][i - 1] + indicator // substitution
      );
    }
  }
  const distance = track[c2.length][c1.length];
  const maxLength = Math.max(c1.length, c2.length);
  return (maxLength - distance) / maxLength;
}

// Default/Initial Mock Data for Flight LH760, LH762, LX146, LX2646
const INITIAL_MOCK_DATA: BaggageRecord[] = [
  {
    id: 'bag-1',
    sno: '1',
    pir: 'FRA_LX_10294',
    name: 'MUELLER HANS',
    originalTag: '0724102943',
    rushTag: 'LX920394',
    flightNo: 'LH760',
    seal: 'S-40192',
    ln: 'L02',
    destination: 'FRA',
    remarks: 'Rush priority passenger connection bag',
    storageRemarks: '',
    status: 'Expected',
    customsStatus: 'Pending',
    disposition: 'Pending',
    registryType: 'Arrival',
    createdAt: new Date().toISOString()
  },
  {
    id: 'bag-2',
    sno: '2',
    pir: 'DEL_LX_49204',
    name: 'SCHMIDT ANNA',
    originalTag: '0724492041',
    rushTag: '',
    flightNo: 'LH762',
    seal: 'S-40195',
    ln: 'L05',
    destination: 'DEL',
    remarks: 'Awaiting customs document pack',
    storageRemarks: 'Rack A-2, Row 3',
    status: 'Received',
    receivedAt: new Date(Date.now() - 4 * 24 * 60 * 60 * 1000).toISOString(), // 4 days ago - triggers 3-day storage alert
    customsStatus: 'Not Cleared',
    customsReason: 'Awaiting documents',
    customsUpdatedAt: new Date(Date.now() - 4 * 24 * 60 * 60 * 1000).toISOString(),
    disposition: 'Storage',
    dispositionLocation: 'LHG Office',
    dispositionUpdatedAt: new Date(Date.now() - 4 * 24 * 60 * 60 * 1000).toISOString(),
    registryType: 'Arrival',
    createdAt: new Date(Date.now() - 4 * 24 * 60 * 60 * 1000).toISOString()
  },
  {
    id: 'bag-3',
    sno: '3',
    pir: 'FRA_LX_88201',
    name: 'ALMEIDA RAHUL',
    originalTag: '0724882015',
    rushTag: 'LX920399',
    flightNo: 'LX146',
    seal: 'S-40221',
    ln: 'L08',
    destination: 'FRA',
    remarks: 'Direct delivery',
    storageRemarks: '',
    status: 'Received',
    receivedAt: new Date(Date.now() - 1 * 24 * 60 * 60 * 1000).toISOString(),
    customsStatus: 'Cleared',
    customsUpdatedAt: new Date(Date.now() - 1 * 24 * 60 * 60 * 1000).toISOString(),
    disposition: 'Delivered',
    dispositionLocation: 'VVM',
    dispositionUpdatedAt: new Date(Date.now() - 1 * 24 * 60 * 60 * 1000).toISOString(),
    registryType: 'Arrival',
    createdAt: new Date(Date.now() - 1 * 24 * 60 * 60 * 1000).toISOString()
  },
  {
    id: 'bag-4',
    sno: '4',
    pir: 'DEL_LX_11029',
    name: 'DUBOIS PIERRE',
    originalTag: '',
    rushTag: 'LX920401',
    flightNo: 'LX2646',
    seal: 'S-40332',
    ln: 'L12',
    destination: 'DEL',
    remarks: 'Re-export candidate if un-cleared',
    storageRemarks: 'Under CWC lock, fragile',
    status: 'Received',
    receivedAt: new Date(Date.now() - 6 * 24 * 60 * 60 * 1000).toISOString(), // 6 days ago - triggers 5-day Urgent Reminder
    customsStatus: 'Not Cleared',
    customsReason: 'Lack of documents',
    customsUpdatedAt: new Date(Date.now() - 6 * 24 * 60 * 60 * 1000).toISOString(),
    disposition: 'Storage',
    dispositionLocation: 'CWC',
    dispositionUpdatedAt: new Date(Date.now() - 6 * 24 * 60 * 60 * 1000).toISOString(),
    registryType: 'Arrival',
    createdAt: new Date(Date.now() - 6 * 24 * 60 * 60 * 1000).toISOString()
  },
  {
    id: 'bag-5',
    sno: '5',
    pir: 'FRA_LX_55212',
    name: 'MEHTA SNEHA',
    originalTag: '0724552120',
    rushTag: '',
    flightNo: 'LH760',
    seal: 'S-40502',
    ln: 'L01',
    destination: 'FRA',
    remarks: 'Needs immediate processing',
    storageRemarks: '',
    status: 'Expected',
    customsStatus: 'Pending',
    disposition: 'Pending',
    registryType: 'Arrival',
    createdAt: new Date().toISOString()
  },
  {
    id: 'bag-6',
    sno: '6',
    pir: 'FRA_LX_77192',
    name: 'KHAN ARMAAN',
    originalTag: '0724771922',
    rushTag: 'LX920555',
    flightNo: 'LH762',
    seal: 'S-40510',
    ln: 'L03',
    destination: 'FRA',
    remarks: 'Disposed old bag - test auto-purge',
    storageRemarks: '',
    status: 'Received',
    receivedAt: new Date(Date.now() - 9 * 24 * 60 * 60 * 1000).toISOString(),
    customsStatus: 'Cleared',
    customsUpdatedAt: new Date(Date.now() - 9 * 24 * 60 * 60 * 1000).toISOString(),
    disposition: 'Delivered',
    dispositionLocation: 'Outlook',
    dispositionUpdatedAt: new Date(Date.now() - 9 * 24 * 60 * 60 * 1000).toISOString(), // 9 days ago - eligible for Auto-Purge
    registryType: 'Arrival',
    createdAt: new Date(Date.now() - 9 * 24 * 60 * 60 * 1000).toISOString()
  },
  {
    id: 'bag-7',
    sno: '7',
    pir: 'DEP_LH_11111',
    name: 'TEST DEPARTURE',
    originalTag: '0724771999',
    rushTag: '',
    flightNo: 'LH762',
    seal: 'S-40599',
    ln: 'L01',
    destination: 'FRA',
    remarks: 'Test departure bag',
    storageRemarks: '',
    status: 'Received',
    receivedAt: new Date(Date.now() - 1 * 24 * 60 * 60 * 1000).toISOString(),
    disposition: 'Awaiting Forwarding',
    dispositionLocation: 'LHG Office',
    dispositionUpdatedAt: new Date(Date.now() - 1 * 24 * 60 * 60 * 1000).toISOString(),
    registryType: 'Departure',
    customsStatus: 'Pending',
    createdAt: new Date(Date.now() - 1 * 24 * 60 * 60 * 1000).toISOString()
  }
];

// Configurable Dispositions & Customs Operations Protocol rules
interface WorkflowFieldConfig {
  name: 'deliveryAgent' | 'storageOption' | 'domesticForwarding' | 'arrivalBelt' | 'handoverOption' | 'warehouseOption' | 'reexportOption';
  label: string;
  type: 'select' | 'text' | 'number';
  options: string[];
  defaultValue: string;
  description?: string;
}

interface ProtocolConfig {
  id: 'Cleared Baggage' | 'Non-Cleared / Other';
  name: string;
  fields: WorkflowFieldConfig[];
}

const PROTOCOL_CONFIGS: ProtocolConfig[] = [
  {
    id: 'Cleared Baggage',
    name: 'Cleared Baggage',
    fields: [
      {
        name: 'deliveryAgent',
        label: 'Delivery Agent',
        type: 'select',
        options: ['VVM', 'Outlook', 'Advik'],
        defaultValue: 'VVM'
      },
      {
        name: 'storageOption',
        label: 'Storage',
        type: 'select',
        options: ['Standard Warehousing – LHG Office'],
        defaultValue: 'Standard Warehousing – LHG Office'
      },
      {
        name: 'domesticForwarding',
        label: 'Domestic Forwarding Airline',
        type: 'select',
        options: ['Air India', 'IndiGo', 'SpiceJet'],
        defaultValue: 'Air India',
        description: 'Selected only if forwarding is required.'
      }
    ]
  },
  {
    id: 'Non-Cleared / Other',
    name: 'Non-Cleared / Other',
    fields: [
      {
        name: 'arrivalBelt',
        label: 'Arrival Belt',
        type: 'select',
        options: ['Arrival Belt 9'],
        defaultValue: 'Arrival Belt 9',
        description: 'Default holding area with queue check.'
      },
      {
        name: 'handoverOption',
        label: 'Handover Option',
        type: 'select',
        options: ['Partner Airlines'],
        defaultValue: 'Partner Airlines',
        description: 'Transfer custody to partner airline.'
      },
      {
        name: 'warehouseOption',
        label: 'Warehouse Option',
        type: 'select',
        options: ['CWC Warehouse'],
        defaultValue: 'CWC Warehouse',
        description: 'Secure central depot storage.'
      },
      {
        name: 'reexportOption',
        label: 'Re-export Option',
        type: 'select',
        options: ['Re-export to Carrier Hub'],
        defaultValue: 'Re-export to Carrier Hub',
        description: 'Bag returned to originating carrier.'
      }
    ]
  }
];

export default function RushBaggageWizard() {
  // Authentication State
  const [user, setUser] = useState<'lh' | 'admin' | null>(null);
  const [usernameInput, setUsernameInput] = useState('');
  const [passwordInput, setPasswordInput] = useState('');
  const [loginError, setLoginError] = useState('');
  const [mounted, setMounted] = useState(false);

  // Mobile-first collapsible navigation state
  const [isMenuOpen, setIsMenuOpen] = useState(false);
  const [activeSection, setActiveSection] = useState<'dashboard' | 'scanner' | 'protocol' | 'registry'>('dashboard');
  const [activeRegistry, setActiveRegistry] = useState<'Arrival' | 'Departure' | 'Combined'>('Combined');
  const [registrationContext, setRegistrationContext] = useState<'arrival' | 'departure'>('arrival');

  // Stable date reference to keep rendering pure
  const [now] = useState(() => Date.now());

  // Primary Data State
  const [baggageList, setBaggageList] = useState<BaggageRecord[]>([]);
  
  // Dashboard Filtering Active State
  const [activeFilter, setActiveFilter] = useState<{
    type: 'all' | 'expected' | 'arrived' | 'non-arrival' | 'cleared' | 'not-cleared' | 'location' | 'alerts';
    value?: string;
  }>({ type: 'all' });

  // Advanced Filters
  const [filterFlights, setFilterFlights] = useState<string[]>([]);
  const [filterDateRange, setFilterDateRange] = useState<{ type: 'all' | 'today' | 'yesterday' | 'last7' | 'last30' | 'thisMonth' | 'prevMonth' | 'custom'; from?: string; to?: string }>({ type: 'all' });
  const [filterRegistry, setFilterRegistry] = useState<'Arrival' | 'Departure' | 'Combined'>('Combined');
  const [filterRemarks, setFilterRemarks] = useState('');

  // Intelligent Scanner States
  const [scannerResult, setScannerResult] = useState<BaggageRecord | null>(null);
  const [scannerResults, setScannerResults] = useState<BaggageRecord[]>([]); // For name searches that return multiples
  const [scannerSearchType, setScannerSearchType] = useState<'none' | 'tag' | 'pir' | 'name'>('none');
  const [scannedQuery, setScannedQuery] = useState('');

  // Quick scan inputs & barcode scanning
  const [scannerInput, setScannerInput] = useState('');
  const [scannerNotification, setScannerNotification] = useState<{
    text: string;
    type: 'success' | 'warning' | 'error';
  } | null>(null);

  // Search & Sorting state
  const [searchTerm, setSearchTerm] = useState('');
  const [flightFilter, setFlightFilter] = useState('ALL');

  // Multi-select state
  const [selectedIds, setSelectedIds] = useState<string[]>([]);

  // Manual bag entry form state
  const [showAddForm, setShowAddForm] = useState(false);
  const [forwardingRequired, setForwardingRequired] = useState(false);
  const [newBag, setNewBag] = useState<Partial<BaggageRecord>>({
    pir: '',
    name: '',
    originalTag: '',
    rushTag: '',
    flightNo: '',
    seal: '',
    ln: '', // Default Locked representation is empty (free text)
    destination: '',
    remarks: '',
    storageRemarks: '',
    receivedAt: new Date().toISOString(),
    status: 'Expected',
    customsStatus: 'Pending',
    disposition: 'Pending',
    weight: undefined,
    damaged: 'N',
    protocol: '', // Mandatory protocol
    deliveryAgent: 'VVM',
    storageOption: 'Standard Warehousing – LHG Office',
    domesticForwarding: 'No Forwarding',
    arrivalBelt: 'Arrival Belt 9',
    handoverOption: 'Partner Airlines',
    warehouseOption: 'CWC Warehouse',
    reexportOption: 'Re-export to Carrier Hub',
    departureForwardingStatus: 'Awaiting Forwarding',
    departureStorageLocation: 'LHG Office',
    leftBehindReason: '',
    bmaStorageLocation: ''
  });

  // Bulk Edit Dialog state
  const [showBulkEdit, setShowBulkEdit] = useState(false);
  const [bulkCustomsStatus, setBulkCustomsStatus] = useState<'Pending' | 'Cleared' | 'Not Cleared' | 'Marked Preventive'>('Cleared');
  const [bulkCustomsReason, setBulkCustomsReason] = useState<BaggageRecord['customsReason']>('');
  const [bulkDisposition, setBulkDisposition] = useState<'Pending' | 'Storage' | 'Delivered' | 'Forwarded' | 'Belt 9' | 'Handover' | 'CWC' | 'Re-export'>('Storage');
  const [bulkLocation, setBulkLocation] = useState<BaggageRecord['dispositionLocation']>('LHG Office');
  
  // Redesign Bulk Edit States
  const [bulkPreventiveReason, setBulkPreventiveReason] = useState('');
  const [bulkHandoverAirline, setBulkHandoverAirline] = useState('');
  const [bulkHandoverDateTime, setBulkHandoverDateTime] = useState('');
  const [bulkWarehouseReference, setBulkWarehouseReference] = useState('');
  const [bulkReexportFlight, setBulkReexportFlight] = useState('');
  const [bulkReexportDestination, setBulkReexportDestination] = useState('');
  const [bulkReexportDate, setBulkReexportDate] = useState('');
  const [bulkDeliveryDateTime, setBulkDeliveryDateTime] = useState('');
  const [bulkForwardingFlightNo, setBulkForwardingFlightNo] = useState('');
  const [bulkForwardingDateNew, setBulkForwardingDateNew] = useState('');
  const [bulkForwardingDestinationNew, setBulkForwardingDestinationNew] = useState('');
  const [bulkStorageRemarks, setBulkStorageRemarks] = useState('');
  const [bulkRemarks, setBulkRemarks] = useState('');

  // Single Edit Dialog state
  const [editingRecord, setEditingRecord] = useState<BaggageRecord | null>(null);

  // Configurable IATA Airline Map state
  const [iataAirlineMap, setIataAirlineMap] = useState<Record<string, string>>(DEFAULT_IATA_AIRLINE_MAP);

  const saveIataAirlineMap = (newMap: Record<string, string>) => {
    setIataAirlineMap(newMap);
    localStorage.setItem('rbw_iata_airline_map', JSON.stringify(newMap));
  };

  // Bulk Add Session States
  const [isBulkMode, setIsBulkMode] = useState(true);
  const [bulkTagsInput, setBulkTagsInput] = useState('');
  const [bulkProtocol, setBulkProtocol] = useState<'Cleared Baggage' | 'Non-Cleared / Other' | ''>('');
  const [bulkClearedAction, setBulkClearedAction] = useState<string>('');
  const [bulkNonClearedAction, setBulkNonClearedAction] = useState<string>('');

  const [bulkDeliveryAgent, setBulkDeliveryAgent] = useState<'VVM' | 'Outlook' | 'Advik' | ''>('VVM');
  const [bulkStorageOption, setBulkStorageOption] = useState<'Standard Warehousing – LHG Office' | ''>('Standard Warehousing – LHG Office');
  const [bulkDomesticForwarding, setBulkDomesticForwarding] = useState<'Air India' | 'IndiGo' | 'SpiceJet' | 'No Forwarding' | ''>('No Forwarding');
  const [bulkForwardingFlight, setBulkForwardingFlight] = useState<string>('');
  const [bulkForwardingDate, setBulkForwardingDate] = useState<string>('');
  const [bulkArrivalBelt, setBulkArrivalBelt] = useState<'Arrival Belt 9' | ''>('Arrival Belt 9');
  const [bulkHandoverOption, setBulkHandoverOption] = useState<'Partner Airlines' | ''>('Partner Airlines');
  const [bulkWarehouseOption, setBulkWarehouseOption] = useState<'CWC Warehouse' | ''>('CWC Warehouse');
  const [bulkReexportOption, setBulkReexportOption] = useState<'Re-export to Carrier Hub' | ''>('Re-export to Carrier Hub');

  // Departure Bulk States
  const [bulkDepartureForwardingStatus, setBulkDepartureForwardingStatus] = useState<'Awaiting Forwarding' | 'Forwarding Sent'>('Awaiting Forwarding');
  const [bulkDepartureStorageLocation, setBulkDepartureStorageLocation] = useState<'LHG Office' | 'BMA' | 'Level 4'>('LHG Office');
  const [bulkLeftBehindReason, setBulkLeftBehindReason] = useState<string>('');
  const [bulkBmaStorageLocation, setBulkBmaStorageLocation] = useState<string>('');
  const [bulkForwardingRemarks, setBulkForwardingRemarks] = useState<string>('');
  const [bulkForwardingDestination, setBulkForwardingDestination] = useState<string>('');
  const [bulkForwardedBy, setBulkForwardedBy] = useState<string>('');

  const [bulkFlightNo, setBulkFlightNo] = useState<string>('');
  const [bulkDestination, setBulkDestination] = useState<string>('');
  const [bulkDamaged, setBulkDamaged] = useState<'Y' | 'N'>('N');
  const [bulkWeight, setBulkWeight] = useState<number | undefined>(undefined);
  const [bulkSeal, setBulkSeal] = useState<string>('');
  const [bulkLn, setBulkLn] = useState<string>('');
  const [bulkStatus, setBulkStatus] = useState<'Expected' | 'Received'>('Expected');

  // Scanner states
  const [isScannerOpen, setIsScannerOpen] = useState(false);
  const [scannerTargetField, setScannerTargetField] = useState<'originalTag' | 'rushTag' | 'bulk' | 'gate'>('originalTag');
  const [continuousScannedTags, setContinuousScannedTags] = useState<string[]>([]);

  // Duplicate resolution state
  const [bulkDuplicatesList, setBulkDuplicatesList] = useState<Array<{
    tag: string;
    isExistingInDb: boolean;
    existingRecord?: BaggageRecord;
    resolution: 'skip' | 'replace' | 'keep';
  }>>([]);
  const [showDuplicatesResolver, setShowDuplicatesResolver] = useState(false);

  // Two-level cascading protocol states
  const [newBagClearedAction, setNewBagClearedAction] = useState<string>('');
  const [newBagNonClearedAction, setNewBagNonClearedAction] = useState<string>('');
  const [editingClearedAction, setEditingClearedAction] = useState<string>('');
  const [editingNonClearedAction, setEditingNonClearedAction] = useState<string>('');

  const handleScanSuccess = (barcode: string) => {
    if (scannerTargetField === 'bulk') {
      setContinuousScannedTags(prev => {
        const next = [...prev, barcode];
        setBulkTagsInput(next.join('\n'));
        return next;
      });
    } else if (scannerTargetField === 'originalTag') {
      setNewBag(prev => ({ ...prev, originalTag: barcode, receivedAt: new Date().toISOString() }));
    } else if (scannerTargetField === 'rushTag') {
      setNewBag(prev => ({ ...prev, rushTag: barcode, receivedAt: new Date().toISOString() }));
    } else if (scannerTargetField === 'gate') {
      handleScanSubmit(undefined, barcode);
      setIsScannerOpen(false);
    }
  };

  const handleFinishContinuous = () => {
    setIsScannerOpen(false);
  };

  const handleOpenEditDialog = (record: BaggageRecord | null) => {
    setEditingRecord(record);
    if (!record) {
      setEditingClearedAction('');
      setEditingNonClearedAction('');
      return;
    }
    if (record.protocol === 'Cleared Baggage') {
      if (record.deliveryAgent) {
        setEditingClearedAction('deliveryAgent');
      } else if (record.storageOption) {
        setEditingClearedAction('storage');
      } else if (record.domesticForwarding && record.domesticForwarding !== 'No Forwarding') {
        setEditingClearedAction('domesticForwarding');
      } else {
        setEditingClearedAction('');
      }
      setEditingNonClearedAction('');
    } else if (record.protocol === 'Non-Cleared / Other') {
      if (record.arrivalBelt) {
        setEditingNonClearedAction('arrivalBelt');
      } else if (record.handoverOption) {
        setEditingNonClearedAction('handover');
      } else if (record.warehouseOption) {
        setEditingNonClearedAction('warehouse');
      } else if (record.reexportOption) {
        setEditingNonClearedAction('reexport');
      } else {
        setEditingNonClearedAction('');
      }
      setEditingClearedAction('');
    } else {
      setEditingClearedAction('');
      setEditingNonClearedAction('');
    }
  };

  // Delete confirmation state for iframe-safe non-blocking overlays
  const [deleteConfirm, setDeleteConfirm] = useState<{
    show: boolean;
    type: 'single' | 'bulk';
    id?: string;
  }>({ show: false, type: 'single' });

  // Import tabs & raw paste state
  const [importTab, setImportTab] = useState<'paste' | 'file'>('paste');
  const [rawPasteText, setRawPasteText] = useState('');
  const [importPreview, setImportPreview] = useState<BaggageRecord[]>([]);
  const [showImportDialog, setShowImportDialog] = useState(false);
  const [showSBHImportDialog, setShowSBHImportDialog] = useState(false);
  const [showSBHInstructions, setShowSBHInstructions] = useState(false);
  const [showAlteaInstructions, setShowAlteaInstructions] = useState(false);
  const [showAlteaPasteWindow, setShowAlteaPasteWindow] = useState(false);
  const [showAlteaPreview, setShowAlteaPreview] = useState(false);
  const [isSBHImport, setIsSBHImport] = useState(false);
  const [isAlteaImport, setIsAlteaImport] = useState(false);
  const [rawAlteaText, setRawAlteaText] = useState('');
  const [alteaPreviewRecords, setAlteaPreviewRecords] = useState<BaggageRecord[]>([]);
  const [selectedAlteaRowIds, setSelectedAlteaRowIds] = useState<Set<string>>(new Set());
  const [alteaSearchQuery, setAlteaSearchQuery] = useState('');
  const [alteaSortField, setAlteaSortField] = useState<keyof BaggageRecord | 'status'>('id');
  const [alteaSortDirection, setAlteaSortDirection] = useState<'asc' | 'desc'>('asc');
  
  // Mobile & OCR States
  const [isMobile, setIsMobile] = useState(false);
  const [isOcrLoading, setIsOcrLoading] = useState(false);
  const [ocrError, setOcrError] = useState<string | null>(null);
  const [ocrTargetMode, setOcrTargetMode] = useState<'altea' | 'table'>('altea');
  const cameraInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    // Detect mobile environment
    const checkMobile = () => {
      setIsMobile(/Android|iPhone|iPad|iPod/i.test(navigator.userAgent));
    };
    checkMobile();
    // Also add listener for window resize in case they toggle device mode
    window.addEventListener('resize', checkMobile);
    return () => window.removeEventListener('resize', checkMobile);
  }, []);

  // Intelligent Importer States
  const [mappingDictionary, setMappingDictionary] = useState<DictionaryEntry[]>(DEFAULT_MAPPING_DICTIONARY);

  const saveMappingDictionary = (newDict: DictionaryEntry[]) => {
    setMappingDictionary(newDict);
    localStorage.setItem('rbw_mapping_dictionary', JSON.stringify(newDict));
  };

  const [importWizardStep, setImportWizardStep] = useState<'upload' | 'mapping' | 'preview' | 'summary'>('upload');
  const [excelHeaders, setExcelHeaders] = useState<string[]>([]);
  const [excelRows, setExcelRows] = useState<any[][]>([]);
  const [filesProcessedCount, setFilesProcessedCount] = useState<number>(0);
  const [columnMappings, setColumnMappings] = useState<Record<number, { systemField: string, confidence: number, matchedBy: 'header' | 'semantic' | 'manual' }>>({});
  const [duplicateMode, setDuplicateMode] = useState<'skip' | 'update' | 'new'>('skip');
  const [importProgress, setImportProgress] = useState<number | null>(null);
  const [importProgressMessage, setImportProgressMessage] = useState<string>('');
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [importSummaryResult, setImportSummaryResult] = useState<{
    totalRows: number,
    imported: number,
    skipped: number,
    duplicates: number,
    invalid: number,
    blankWeights: number,
    invalidWeights: number,
    recognizedLocks: number,
    unrecognizedLocks: number,
    warnings: string[],
    filesProcessed?: number
  } | null>(null);

  const [newAliasField, setNewAliasField] = useState<string>('');
  const [newAliasValue, setNewAliasValue] = useState<string>('');
  const [showDictionaryEditor, setShowDictionaryEditor] = useState(false);

  // Lock Mapping Dictionary States
  const [lockDictionary, setLockDictionary] = useState<Record<string, string>>(DEFAULT_LOCK_DICTIONARY);

  const saveLockDictionary = (newDict: Record<string, string>) => {
    setLockDictionary(newDict);
    localStorage.setItem('rbw_lock_dictionary', JSON.stringify(newDict));
  };

  const [showLockDictionaryEditor, setShowLockDictionaryEditor] = useState(false);
  const [newLockAbbr, setNewLockAbbr] = useState('');
  const [newLockExpanded, setNewLockExpanded] = useState('');

  // System status flags
  const [purgeAlertCount, setPurgeAlertCount] = useState(0);

  // Load from local storage
  useEffect(() => {
    setTimeout(() => {
      setMounted(true);
    }, 0);
    const storedUser = localStorage.getItem('rbw_logged_in_user');
    if (storedUser === 'lh' || storedUser === 'admin') {
      // eslint-disable-next-line react-hooks/set-state-in-effect
      setUser(storedUser);
    }
    
    const storedData = localStorage.getItem('rbw_baggage_list');
    if (storedData) {
      try {
        const parsed = JSON.parse(storedData) as BaggageRecord[];
        // Data Migration: Ensure receivedAt is present
        const migrated = parsed.map(b => ({
          ...b,
          receivedAt: b.receivedAt || b.createdAt
        }));
        setBaggageList(migrated);
      } catch (e) {
        setBaggageList(INITIAL_MOCK_DATA);
      }
    } else {
      // Default initial records
      setBaggageList(INITIAL_MOCK_DATA);
      localStorage.setItem('rbw_baggage_list', JSON.stringify(INITIAL_MOCK_DATA));
    }

    // Load additional dictionaries to prevent hydration mismatch
    const savedIata = localStorage.getItem('rbw_iata_airline_map');
    if (savedIata) {
      try {
        setIataAirlineMap(JSON.parse(savedIata));
      } catch (e) {}
    }

    const storedDict = localStorage.getItem('rbw_mapping_dictionary');
    if (storedDict) {
      try {
        setMappingDictionary(JSON.parse(storedDict));
      } catch (e) {}
    }

    const storedLock = localStorage.getItem('rbw_lock_dictionary');
    if (storedLock) {
      try {
        setLockDictionary(JSON.parse(storedLock));
      } catch (e) {}
    }
  }, []);

  // Save to local storage helper
  const saveBaggageData = (data: BaggageRecord[]) => {
    setBaggageList(data);
    localStorage.setItem('rbw_baggage_list', JSON.stringify(data));
  };

  // Perform Auto-Purge check on Load or trigger
  const runAutoPurge = () => {
    const purgeTimeLimit = 7 * 24 * 60 * 60 * 1000; // 7 days in ms
    const now = Date.now();
    
    // Disposed status list: 'Delivered', 'Forwarded', 'Handover', 'Re-export'
    const disposedDispositions = ['Delivered', 'Forwarded', 'Handover', 'Re-export'];
    
    const initialCount = baggageList.length;
    const filtered = baggageList.filter(bag => {
      if (disposedDispositions.includes(bag.disposition) && bag.dispositionUpdatedAt) {
        const timeSinceDisposal = now - new Date(bag.dispositionUpdatedAt).getTime();
        return timeSinceDisposal < purgeTimeLimit;
      }
      return true;
    });

    const purgedCount = initialCount - filtered.length;
    if (purgedCount > 0) {
      saveBaggageData(filtered);
      setPurgeAlertCount(purgedCount);
      setTimeout(() => setPurgeAlertCount(0), 10000); // Clear alert banner after 10s
    }
  };

  // Run auto-purge on initialization once list is ready
  useEffect(() => {
    if (baggageList.length > 0) {
      // eslint-disable-next-line react-hooks/set-state-in-effect
      runAutoPurge();
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [baggageList.length]);

  // Login handler
  const handleLogin = (e: React.FormEvent) => {
    e.preventDefault();
    const u = usernameInput.trim().toLowerCase();
    const p = passwordInput;

    if (u === 'lh' && p === 'welcome') {
      setUser('lh');
      localStorage.setItem('rbw_logged_in_user', 'lh');
      setLoginError('');
    } else if (u === 'admin' && p === 'Admin220!') {
      setUser('admin');
      localStorage.setItem('rbw_logged_in_user', 'admin');
      setLoginError('');
    } else {
      setLoginError('Invalid Username or Password. Please try again.');
    }
  };

  // Logout handler
  const handleLogout = () => {
    setUser(null);
    localStorage.removeItem('rbw_logged_in_user');
    setUsernameInput('');
    setPasswordInput('');
  };

  // Calculate Days in Storage helper
  const getDaysInStorage = React.useCallback((record: BaggageRecord): number => {
    if (!record.receivedAt || record.status !== 'Received') return 0;
    // Only count days if currently stored in an active storage location
    const storageLocations = ['Belt 9', 'LHG Office', 'BMA', 'Level 4 Checks', 'CWC'];
    if (!record.dispositionLocation || !storageLocations.includes(record.dispositionLocation)) {
      return 0;
    }
    
    const entryDate = new Date(record.receivedAt);
    const diffTime = Math.abs(now - entryDate.getTime());
    return Math.floor(diffTime / (1000 * 60 * 60 * 24));
  }, [now]);

  // Compute stats for Dashboard
  const stats = useMemo(() => {
    const filteredBaggageList = baggageList.filter(item => activeRegistry === 'Combined' || item.registryType === activeRegistry);
    const totalExpected = filteredBaggageList.length;
    const arrived = filteredBaggageList.filter(b => b.status === 'Received').length;
    const nonArrivals = filteredBaggageList.filter(b => b.status === 'Expected').length;
    const cleared = filteredBaggageList.filter(b => b.status === 'Received' && b.customsStatus === 'Cleared').length;
    const notCleared = filteredBaggageList.filter(b => b.status === 'Received' && (b.customsStatus === 'Not Cleared' || b.customsStatus === 'Marked Preventive')).length;

    // Location counters
    const belt9 = filteredBaggageList.filter(b => b.status === 'Received' && b.dispositionLocation === 'Belt 9').length;
    const lhgOffice = filteredBaggageList.filter(b => b.status === 'Received' && b.dispositionLocation === 'LHG Office').length;
    const bma = filteredBaggageList.filter(b => b.status === 'Received' && b.dispositionLocation === 'BMA').length;
    const level4Checks = filteredBaggageList.filter(b => b.status === 'Received' && b.dispositionLocation === 'Level 4 Checks').length;
    const cwc = filteredBaggageList.filter(b => b.status === 'Received' && b.dispositionLocation === 'CWC').length;

    // Total alerts count (3 days and 5 days)
    const alert3Days = filteredBaggageList.filter(b => getDaysInStorage(b) >= 3 && getDaysInStorage(b) < 5).length;
    const alert5Days = filteredBaggageList.filter(b => getDaysInStorage(b) >= 5).length;

    return {
      totalExpected,
      arrived,
      nonArrivals,
      cleared,
      notCleared,
      locations: {
        'Belt 9': belt9,
        'LHG Office': lhgOffice,
        'BMA': bma,
        'Level 4 Checks': level4Checks,
        'CWC': cwc
      },
      alert3Days,
      alert5Days
    };
  }, [baggageList, activeRegistry, getDaysInStorage]);

  // Scanner/Reconciliation Matcher
  const handleScanSubmit = (e?: React.FormEvent, overrideQuery?: string) => {
    if (e) e.preventDefault();
    const query = (overrideQuery || scannerInput).trim().toUpperCase();
    if (!query) return;

    setScannerNotification(null);
    setScannerResult(null);
    setScannerResults([]);
    setScannedQuery(query);

    // 1. Detect Search Type
    const isLikelyTag = /^\d{10}$/.test(query) || /^[A-Z]{2}\d+$/.test(query) || /^\d+$/.test(query);
    const isLikelyPIR = /^[A-Z]{3}[A-Z]{2}\d{5}$/.test(query); // e.g. DELEK12345
    
    let type: 'tag' | 'pir' | 'name' = 'name';
    if (isLikelyTag) type = 'tag';
    else if (isLikelyPIR) type = 'pir';

    setScannerSearchType(type);

    if (type === 'tag') {
      const canonicalTag = getCanonicalTag(query, iataAirlineMap);
      const match = baggageList.find(b => 
        (b.originalTag && matchTag(b.originalTag, query, iataAirlineMap)) ||
        (b.rushTag && matchTag(b.rushTag, query, iataAirlineMap))
      );

      if (match) {
        setScannerResult(match);
        setScannerNotification({
          text: `MATCH FOUND: Baggage Tag ${canonicalTag} is already registered.`,
          type: 'success'
        });
      } else {
        setScannerNotification({
          text: `No matching baggage record found. Would you like to register this baggage?`,
          type: 'warning'
        });
      }
    } else if (type === 'pir') {
      const match = baggageList.find(b => b.pir && b.pir.toUpperCase() === query);
      if (match) {
        setScannerResult(match);
      } else {
        setScannerNotification({
          text: `NO RECORD FOUND: No baggage record for PIR '${query}'. Please register using a Tag number.`,
          type: 'error'
        });
      }
    } else {
      const matches = baggageList.filter(b => b.name.toUpperCase().includes(query));
      if (matches.length > 0) {
        setScannerResults(matches);
      } else {
        setScannerNotification({
          text: `NO RECORD FOUND: No results for '${query}'. Please register via Tag number.`,
          type: 'error'
        });
      }
    }
    setScannerInput('');
  };

  // Manual Addition
  const handleAddBagSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newBag.flightNo) {
      alert('Flight No is required');
      return;
    }

    const hasOriginalTag = !!(newBag.originalTag && newBag.originalTag.trim());
    const hasRushTag = !!(newBag.rushTag && newBag.rushTag.trim());
    const hasName = !!(newBag.name && newBag.name.trim());
    const hasPir = !!(newBag.pir && newBag.pir.trim());

    if (!hasOriginalTag && !hasRushTag && !hasName && !hasPir) {
      alert('Please enter at least one identifying field:\nOriginal Tag, Rush Tag, Passenger Name or PIR Number.');
      return;
    }

    // Weight validation: If provided, must be a positive number
    if (newBag.weight !== undefined && (isNaN(Number(newBag.weight)) || Number(newBag.weight) < 0)) {
      alert('Weight (kg) must be a positive number');
      return;
    }

    // Assign a default protocol and default sub-actions to ensure save succeeds without blocks
    const finalProtocol = newBag.protocol || 'Cleared Baggage';
    const finalClearedAction = newBagClearedAction || 'deliveryAgent';
    const finalNonClearedAction = newBagNonClearedAction || 'arrivalBelt';

    const canonicalOriginalTag = getCanonicalTag(newBag.originalTag || '', iataAirlineMap);
    const canonicalRushTag = getCanonicalTag(newBag.rushTag || '', iataAirlineMap);

    const duplicate = baggageList.find(b => 
      (canonicalOriginalTag && b.originalTag === canonicalOriginalTag) ||
      (canonicalRushTag && b.rushTag === canonicalRushTag)
    );

    if (duplicate) {
      if (confirm(`A duplicate record already exists (PIR: ${duplicate.pir}, Name: ${duplicate.name}). Delete the existing record and add this one instead?`)) {
        setBaggageList(prev => prev.filter(b => b.id !== duplicate.id));
      } else {
        return;
      }
    }

    const createdRecord: BaggageRecord = {
      id: `bag-${Date.now()}`,
      sno: (baggageList.length + 1).toString(), // Auto-assign sno, remove from form
      pir: (newBag.pir || '').toUpperCase() || 'UNKNOWN PIR',
      name: (newBag.name || '').toUpperCase() || 'UNKNOWN PASSENGER',
      originalTag: canonicalOriginalTag,
      rushTag: canonicalRushTag,
      flightNo: newBag.flightNo || '',
      seal: (newBag.seal || '').toUpperCase(),
      ln: newBag.ln || '', // Free text for Locked indicator (such as PAD, CL, Y, N)
      destination: (newBag.destination || '').toUpperCase(),
      remarks: registrationContext === 'arrival' ? (newBag.remarks || 'Additional manually registered bag') : '',
      storageRemarks: newBag.storageRemarks || '',
      status: newBag.status as 'Expected' | 'Received',
      receivedAt: newBag.receivedAt || (newBag.status === 'Received' ? new Date().toISOString() : undefined),
      customsStatus: (registrationContext === 'arrival' ? (newBag.customsStatus || 'Pending') : 'Pending') as BaggageRecord['customsStatus'],
      customsReason: (registrationContext === 'arrival' && newBag.customsStatus === 'Not Cleared' ? newBag.customsReason : '') as BaggageRecord['customsReason'],
      customsUpdatedAt: newBag.status === 'Received' ? new Date().toISOString() : undefined,
      disposition: registrationContext === 'arrival' ? (newBag.status === 'Received' ? 'Storage' : 'Pending') : (newBag.departureForwardingStatus === 'Forwarding Sent' ? 'Forwarded' : 'Awaiting Forwarding'),
      dispositionLocation: registrationContext === 'arrival' ? (newBag.status === 'Received' ? 'LHG Office' : '') : (newBag.departureStorageLocation || ''),
      dispositionUpdatedAt: newBag.status === 'Received' ? new Date().toISOString() : undefined,
      registryType: registrationContext === 'arrival' ? 'Arrival' : 'Departure',
      createdAt: new Date().toISOString(),

      // Departure Specific
      departureForwardingStatus: registrationContext === 'departure' ? newBag.departureForwardingStatus : undefined,
      departureStorageLocation: registrationContext === 'departure' ? newBag.departureStorageLocation : undefined,
      leftBehindReason: registrationContext === 'departure' ? newBag.leftBehindReason : undefined,
      bmaStorageLocation: (registrationContext === 'departure' && newBag.departureStorageLocation === 'BMA') ? newBag.bmaStorageLocation : undefined,
      forwardingFlightNo: (registrationContext === 'departure' && newBag.departureForwardingStatus === 'Forwarding Sent') ? newBag.forwardingFlightNo : undefined,
      forwardingDestination: (registrationContext === 'departure' && newBag.departureForwardingStatus === 'Forwarding Sent') ? newBag.forwardingDestination : undefined,
      forwardingDate: (registrationContext === 'departure' && newBag.departureForwardingStatus === 'Forwarding Sent') ? newBag.forwardingDate : undefined,
      forwardedBy: (registrationContext === 'departure' && newBag.departureForwardingStatus === 'Forwarding Sent') ? newBag.forwardedBy : undefined,
      forwardingRemarks: (registrationContext === 'departure' && newBag.departureForwardingStatus === 'Forwarding Sent') ? newBag.forwardingRemarks : undefined,

      // New properties
      weight: newBag.weight !== undefined && !isNaN(Number(newBag.weight)) ? Number(newBag.weight) : undefined,
      damaged: (newBag.damaged || 'N') as 'Y' | 'N',
      protocol: registrationContext === 'arrival' ? (finalProtocol as 'Cleared Baggage' | 'Non-Cleared / Other') : undefined,
      deliveryAgent: (registrationContext === 'arrival' && finalProtocol === 'Cleared Baggage' && finalClearedAction === 'deliveryAgent') ? (newBag.deliveryAgent || 'VVM') as any : undefined,
      storageOption: (registrationContext === 'arrival' && finalProtocol === 'Cleared Baggage' && finalClearedAction === 'storage') ? (newBag.storageOption || 'Standard Warehousing – LHG Office') as any : undefined,
      domesticForwarding: (registrationContext === 'arrival' && finalProtocol === 'Cleared Baggage' && finalClearedAction === 'domesticForwarding') ? (newBag.domesticForwarding || 'No Forwarding') as any : undefined,
      arrivalBelt: (registrationContext === 'arrival' && finalProtocol === 'Non-Cleared / Other' && finalNonClearedAction === 'arrivalBelt') ? (newBag.arrivalBelt || 'Arrival Belt 9') as any : undefined,
      handoverOption: (registrationContext === 'arrival' && finalProtocol === 'Non-Cleared / Other' && finalNonClearedAction === 'handover') ? (newBag.handoverOption || 'Partner Airlines') as any : undefined,
      warehouseOption: (finalProtocol === 'Non-Cleared / Other' && finalNonClearedAction === 'warehouse') ? (newBag.warehouseOption || 'CWC Warehouse') as any : undefined,
      reexportOption: (finalProtocol === 'Non-Cleared / Other' && finalNonClearedAction === 'reexport') ? (newBag.reexportOption || 'Re-export to Carrier Hub') as any : undefined
    };

    saveBaggageData([createdRecord, ...baggageList]);
    setShowAddForm(false);
    // Reset manual form
    setNewBag({
      pir: '',
      name: '',
      originalTag: '',
      rushTag: '',
      flightNo: 'LH760',
      seal: '',
      ln: 'No',
      destination: '',
      remarks: '',
      storageRemarks: '',
      status: 'Expected',
      customsStatus: 'Pending',
      disposition: 'Pending',
      weight: undefined,
      damaged: 'N',
      protocol: '',
      deliveryAgent: 'VVM',
      storageOption: 'Standard Warehousing – LHG Office',
      domesticForwarding: 'No Forwarding',
      arrivalBelt: 'Arrival Belt 9',
      handoverOption: 'Partner Airlines',
      warehouseOption: 'CWC Warehouse',
      reexportOption: 'Re-export to Carrier Hub'
    });
    setNewBagClearedAction('');
    setNewBagNonClearedAction('');
    setForwardingRequired(false);
  };

  const findAndResolveDuplicates = () => {
    const records = baggageList;
    let found = false;
    for (let i = 0; i < records.length; i++) {
        for (let j = i + 1; j < records.length; j++) {
            const isDuplicate = 
              (records[i].originalTag && records[i].originalTag === records[j].originalTag) ||
              (records[i].rushTag && records[i].rushTag === records[j].rushTag);
            
            if (isDuplicate) {
                found = true;
                if (confirm(`Duplicate found:\n1. PIR: ${records[i].pir}, Name: ${records[i].name}, Tag: ${records[i].originalTag || records[i].rushTag}\n2. PIR: ${records[j].pir}, Name: ${records[j].name}, Tag: ${records[j].originalTag || records[j].rushTag}\n\nDelete the first one?`)) {
                    setBaggageList(prev => prev.filter(b => b.id !== records[i].id));
                    return; 
                } else if (confirm(`Delete the second one?`)) {
                    setBaggageList(prev => prev.filter(b => b.id !== records[j].id));
                    return;
                }
            }
        }
    }
    if (!found) {
        alert('No duplicates found!');
    }
  }

  // Parses a string with potential delimiters (spaces, commas, tabs, newlines) into individual tags
  const parseBulkInput = (text: string): string[] => {
    const rawTokens = text.split(/[\s,\t\r\n]+/);
    return rawTokens.map(t => t.trim()).filter(t => t.length > 0);
  };

  const handleProcessBulkAdd = (e: React.FormEvent) => {
    e.preventDefault();
    const tags = parseBulkInput(bulkTagsInput);
    if (tags.length === 0) {
      alert('Please enter or scan at least one baggage tag number.');
      return;
    }

    // Check for duplicates
    const detectedDuplicates: Array<{
      tag: string;
      isExistingInDb: boolean;
      existingRecord?: BaggageRecord;
      resolution: 'skip' | 'replace' | 'keep';
    }> = [];

    const uniqueTagsInSession = new Set<string>();

    tags.forEach(tag => {
      const canonicalTag = getCanonicalTag(tag, iataAirlineMap);
      
      const isSessionDup = uniqueTagsInSession.has(canonicalTag);
      uniqueTagsInSession.add(canonicalTag);

      const existingRecord = baggageList.find(b => 
        (b.originalTag && matchTag(b.originalTag, canonicalTag, iataAirlineMap)) ||
        (b.rushTag && matchTag(b.rushTag, canonicalTag, iataAirlineMap))
      );

      if (isSessionDup || existingRecord) {
        if (!detectedDuplicates.some(d => d.tag === canonicalTag)) {
          detectedDuplicates.push({
            tag: canonicalTag,
            isExistingInDb: !!existingRecord,
            existingRecord,
            resolution: 'skip'
          });
        }
      }
    });

    if (detectedDuplicates.length > 0) {
      setBulkDuplicatesList(detectedDuplicates);
      setShowDuplicatesResolver(true);
    } else {
      saveBulkRecords(tags, []);
    }
  };

  const saveBulkRecords = (allTags: string[], resolvedDuplicates: typeof bulkDuplicatesList) => {
    const finalRecords: BaggageRecord[] = [];
    const processedCanonicalTags = new Set<string>();

    const tagsToSkip = new Set(resolvedDuplicates.filter(d => d.resolution === 'skip').map(d => d.tag));
    const tagsToReplace = resolvedDuplicates.filter(d => d.resolution === 'replace');
    const tagsToKeep = new Set(resolvedDuplicates.filter(d => d.resolution === 'keep').map(d => d.tag));

    allTags.forEach(tag => {
      const canonical = getCanonicalTag(tag, iataAirlineMap);
      
      if (tagsToSkip.has(canonical)) {
        return;
      }

      const isReplace = tagsToReplace.some(d => d.tag === canonical);
      if (isReplace && !tagsToKeep.has(canonical)) {
        return;
      }

      if (processedCanonicalTags.has(canonical) && !tagsToKeep.has(canonical)) {
        return;
      }
      processedCanonicalTags.add(canonical);

      const finalProtocol = registrationContext === 'arrival' ? (bulkProtocol || 'Cleared Baggage') : undefined;
      const isReceived = bulkStatus === 'Received';
      
      const record: BaggageRecord = {
        id: `bag-${Date.now()}-${Math.floor(Math.random() * 100000)}`,
        sno: (baggageList.length + finalRecords.length + 1).toString(),
        pir: 'UNKNOWN PIR',
        name: 'UNKNOWN PASSENGER',
        originalTag: canonical,
        rushTag: '',
        flightNo: bulkFlightNo,
        seal: bulkSeal.toUpperCase(),
        ln: bulkLn,
        destination: registrationContext === 'arrival' ? bulkDestination.toUpperCase() : bulkForwardingDestination.toUpperCase(),
        remarks: registrationContext === 'arrival' ? `Bulk registered during session (${new Date().toLocaleDateString()})` : '',
        storageRemarks: registrationContext === 'arrival' ? '' : bulkForwardingRemarks,
        status: bulkStatus,
        receivedAt: newBag.receivedAt || (isReceived ? new Date().toISOString() : undefined),
        customsStatus: 'Pending',
        customsUpdatedAt: isReceived ? new Date().toISOString() : undefined,
        disposition: registrationContext === 'arrival' ? (isReceived ? 'Storage' : 'Pending') : (bulkDepartureForwardingStatus === 'Forwarding Sent' ? 'Forwarded' : 'Awaiting Forwarding'),
        dispositionLocation: registrationContext === 'arrival' ? (isReceived ? 'LHG Office' : '') : bulkDepartureStorageLocation,
        dispositionUpdatedAt: isReceived ? new Date().toISOString() : undefined,
        createdAt: new Date().toISOString(),

        // Departure Specific
        departureForwardingStatus: registrationContext === 'departure' ? bulkDepartureForwardingStatus : undefined,
        departureStorageLocation: registrationContext === 'departure' ? bulkDepartureStorageLocation : undefined,
        leftBehindReason: registrationContext === 'departure' ? bulkLeftBehindReason : undefined,
        bmaStorageLocation: (registrationContext === 'departure' && bulkDepartureStorageLocation === 'BMA') ? bulkBmaStorageLocation : undefined,
        forwardingFlightNo: (registrationContext === 'departure' && bulkDepartureForwardingStatus === 'Forwarding Sent') ? bulkForwardingFlight : undefined,
        forwardingDestination: (registrationContext === 'departure' && bulkDepartureForwardingStatus === 'Forwarding Sent') ? bulkForwardingDestination : undefined,
        forwardingDate: (registrationContext === 'departure' && bulkDepartureForwardingStatus === 'Forwarding Sent') ? bulkForwardingDate : undefined,
        forwardedBy: (registrationContext === 'departure' && bulkDepartureForwardingStatus === 'Forwarding Sent') ? bulkForwardedBy : undefined,
        forwardingRemarks: (registrationContext === 'departure' && bulkDepartureForwardingStatus === 'Forwarding Sent') ? bulkForwardingRemarks : undefined,

        weight: bulkWeight,
        damaged: bulkDamaged,
        protocol: finalProtocol as any,
        deliveryAgent: (registrationContext === 'arrival' && finalProtocol === 'Cleared Baggage' && bulkClearedAction === 'deliveryAgent') ? bulkDeliveryAgent as any : undefined,
        storageOption: (registrationContext === 'arrival' && finalProtocol === 'Cleared Baggage' && bulkClearedAction === 'storage') ? bulkStorageOption as any : undefined,
        domesticForwarding: (registrationContext === 'arrival' && finalProtocol === 'Cleared Baggage' && bulkClearedAction === 'domesticForwarding') ? bulkDomesticForwarding as any : undefined,
        arrivalBelt: (registrationContext === 'arrival' && finalProtocol === 'Non-Cleared / Other' && bulkNonClearedAction === 'arrivalBelt') ? bulkArrivalBelt as any : undefined,
        handoverOption: (registrationContext === 'arrival' && finalProtocol === 'Non-Cleared / Other' && bulkNonClearedAction === 'handover') ? bulkHandoverOption as any : undefined,
        warehouseOption: (registrationContext === 'arrival' && finalProtocol === 'Non-Cleared / Other' && bulkNonClearedAction === 'warehouse') ? bulkWarehouseOption as any : undefined,
        reexportOption: (registrationContext === 'arrival' && finalProtocol === 'Non-Cleared / Other' && bulkNonClearedAction === 'reexport') ? bulkReexportOption as any : undefined,
        registryType: registrationContext === 'arrival' ? 'Arrival' : 'Departure' as any
      };

      finalRecords.push(record);
    });

    const updatedDbList = baggageList.map(item => {
      const matchReplacement = tagsToReplace.find(d => 
        (item.originalTag && matchTag(item.originalTag, d.tag, iataAirlineMap)) ||
        (item.rushTag && matchTag(item.rushTag, d.tag, iataAirlineMap))
      );

      if (matchReplacement) {
        const finalProtocol = registrationContext === 'arrival' ? (bulkProtocol || 'Cleared Baggage') : undefined;
        const isReceived = bulkStatus === 'Received';

        return {
          ...item,
          flightNo: bulkFlightNo,
          seal: bulkSeal.toUpperCase() || item.seal,
          ln: bulkLn || item.ln,
          destination: registrationContext === 'arrival' ? (bulkDestination.toUpperCase() || item.destination) : (bulkForwardingDestination.toUpperCase() || item.destination),
          status: bulkStatus,
          receivedAt: isReceived ? (item.receivedAt || new Date().toISOString()) : item.receivedAt,
          weight: bulkWeight !== undefined ? bulkWeight : item.weight,
          damaged: bulkDamaged !== undefined ? bulkDamaged : item.damaged,
          protocol: finalProtocol as any,

          // Departure Specific
          departureForwardingStatus: registrationContext === 'departure' ? bulkDepartureForwardingStatus : item.departureForwardingStatus,
          departureStorageLocation: registrationContext === 'departure' ? bulkDepartureStorageLocation : item.departureStorageLocation,
          leftBehindReason: registrationContext === 'departure' ? bulkLeftBehindReason : item.leftBehindReason,
          bmaStorageLocation: (registrationContext === 'departure' && bulkDepartureStorageLocation === 'BMA') ? bulkBmaStorageLocation : item.bmaStorageLocation,
          forwardingFlightNo: (registrationContext === 'departure' && bulkDepartureForwardingStatus === 'Forwarding Sent') ? bulkForwardingFlight : item.forwardingFlightNo,
          forwardingDestination: (registrationContext === 'departure' && bulkDepartureForwardingStatus === 'Forwarding Sent') ? bulkForwardingDestination : item.forwardingDestination,
          forwardingDate: (registrationContext === 'departure' && bulkDepartureForwardingStatus === 'Forwarding Sent') ? bulkForwardingDate : item.forwardingDate,
          forwardedBy: (registrationContext === 'departure' && bulkDepartureForwardingStatus === 'Forwarding Sent') ? bulkForwardedBy : item.forwardedBy,
          forwardingRemarks: (registrationContext === 'departure' && bulkDepartureForwardingStatus === 'Forwarding Sent') ? bulkForwardingRemarks : item.forwardingRemarks,

          deliveryAgent: (registrationContext === 'arrival' && finalProtocol === 'Cleared Baggage' && bulkClearedAction === 'deliveryAgent') ? bulkDeliveryAgent as any : undefined,
          storageOption: (registrationContext === 'arrival' && finalProtocol === 'Cleared Baggage' && bulkClearedAction === 'storage') ? bulkStorageOption as any : undefined,
          domesticForwarding: (registrationContext === 'arrival' && finalProtocol === 'Cleared Baggage' && bulkClearedAction === 'domesticForwarding') ? bulkDomesticForwarding as any : undefined,
          arrivalBelt: (registrationContext === 'arrival' && finalProtocol === 'Non-Cleared / Other' && bulkNonClearedAction === 'arrivalBelt') ? bulkArrivalBelt as any : undefined,
          handoverOption: (registrationContext === 'arrival' && finalProtocol === 'Non-Cleared / Other' && bulkNonClearedAction === 'handover') ? bulkHandoverOption as any : undefined,
          warehouseOption: (registrationContext === 'arrival' && finalProtocol === 'Non-Cleared / Other' && bulkNonClearedAction === 'warehouse') ? bulkWarehouseOption as any : undefined,
          reexportOption: (registrationContext === 'arrival' && finalProtocol === 'Non-Cleared / Other' && bulkNonClearedAction === 'reexport') ? bulkReexportOption as any : undefined,
          registryType: registrationContext === 'arrival' ? 'Arrival' : 'Departure' as any
        };
      }
      return item;
    });

    saveBaggageData([...finalRecords, ...updatedDbList]);
    
    setShowAddForm(false);
    setShowDuplicatesResolver(false);
    setBulkTagsInput('');
    setContinuousScannedTags([]);
    setBulkDuplicatesList([]);
    alert(`Successfully registered ${finalRecords.length} new records and updated ${tagsToReplace.length} existing records!`);
  };

  // Single record editing update
  const handleEditSaveSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingRecord) return;

    // Protocol and context-specific validation
    if (editingRecord.registryType === 'Departure') {
      if (!editingRecord.departureForwardingStatus) {
        alert('Forwarding Status is required');
        return;
      }
      if (!editingRecord.departureStorageLocation) {
        alert('Storage Location is required');
        return;
      }
      if (editingRecord.departureForwardingStatus === 'Forwarding Sent') {
        if (!editingRecord.forwardingFlightNo || !editingRecord.forwardingDate || !editingRecord.forwardingDestination || !editingRecord.forwardedBy) {
          alert('All forwarding details are mandatory when status is Forwarding Sent');
          return;
        }
      }
      if (editingRecord.departureStorageLocation === 'BMA' && !editingRecord.bmaStorageLocation) {
        alert('BMA Storage Location / Remarks is required');
        return;
      }
    } else {
      // Arrival Specific Validation
      if (!editingRecord.customsStatus || editingRecord.customsStatus === 'Pending') {
        alert('Customs Status is required');
        return;
      }

      if (editingRecord.customsStatus === 'Not Cleared' || editingRecord.customsStatus === 'Marked Preventive') {
        if (!editingRecord.preventiveReason) {
          alert('Preventive Reason is required');
          return;
        }
        if (!editingRecord.disposition || editingRecord.disposition === 'Pending') {
          alert('Disposition is required');
          return;
        }

        if (editingRecord.disposition === 'Handover') {
          if (!editingRecord.handoverAirline) { alert('Partner Airline is required'); return; }
          if (!editingRecord.handoverDateTime) { alert('Handover Date & Time is required'); return; }
        }
        if (editingRecord.disposition === 'CWC') {
          if (!editingRecord.warehouseReference) { alert('Warehouse Reference is required'); return; }
        }
        if (editingRecord.disposition === 'Re-export') {
          if (!editingRecord.reexportFlight) { alert('Re-export Flight is required'); return; }
          if (!editingRecord.reexportDestination) { alert('Destination is required'); return; }
          if (!editingRecord.reexportDate) { alert('Re-export Date is required'); return; }
        }
      } else if (editingRecord.customsStatus === 'Cleared') {
        if (!editingRecord.disposition || editingRecord.disposition === 'Pending') {
          alert('Disposition is required');
          return;
        }

        if (editingRecord.disposition === 'Delivered') {
          if (!editingRecord.deliveryAgent) { alert('Delivery Agent is required'); return; }
          if (!editingRecord.deliveryDateTime) { alert('Delivery Date & Time is required'); return; }
        }
        if (editingRecord.disposition === 'Forwarded') {
          if (!editingRecord.domesticForwarding) { alert('Forward Via is required'); return; }
          if (!editingRecord.forwardingFlightNo) { alert('Forwarding Flight Number is required'); return; }
          if (!editingRecord.forwardingDate) { alert('Forwarding Date is required'); return; }
          if (!editingRecord.forwardingDestination) { alert('Destination is required'); return; }
        }
      }
    }

    const getLegacyFieldsFromProtocol = (
      record: BaggageRecord
    ) => {
      if (record.registryType === 'Departure') {
        return {
          status: record.status || 'Received',
          customsStatus: 'Pending' as const,
          disposition: (record.departureForwardingStatus === 'Forwarding Sent' ? 'Forwarded' : 'Awaiting Forwarding') as any,
          dispositionLocation: record.departureStorageLocation || 'LHG Office'
        };
      }

      // Arrival Redesign Mapping
      let dispositionLocation: any = '';
      if (record.disposition === 'Delivered') {
        dispositionLocation = record.deliveryAgent || 'VVM';
      } else if (record.disposition === 'Storage') {
        dispositionLocation = 'LHG Office';
      } else if (record.disposition === 'Forwarded') {
        const df = record.domesticForwarding || 'Air India';
        dispositionLocation = df === 'IndiGo' ? 'Indigo' : df === 'SpiceJet' ? 'Spice Jet' : 'Air India';
      } else if (record.disposition === 'Belt 9') {
        dispositionLocation = 'Belt 9';
      } else if (record.disposition === 'Handover') {
        dispositionLocation = record.handoverAirline || 'Other Airline';
      } else if (record.disposition === 'CWC') {
        dispositionLocation = 'CWC';
      } else if (record.disposition === 'Re-export') {
        dispositionLocation = 'Hub Re-export';
      }

      return {
        status: 'Received' as const,
        customsStatus: record.customsStatus,
        disposition: record.disposition,
        dispositionLocation
      };
    };

    const updatedList = baggageList.map(item => {
      if (item.id === editingRecord.id) {
        const legacyFields = getLegacyFieldsFromProtocol(editingRecord);

        const wasReceived = item.status === 'Received';
        const isReceived = legacyFields.status === 'Received';
        
        let receivedAt = editingRecord.receivedAt || item.receivedAt;
        if (!wasReceived && isReceived) {
          receivedAt = new Date().toISOString();
        } else if (!isReceived && editingRecord.registryType !== 'Departure') {
          receivedAt = undefined;
        }

        let dispositionUpdatedAt = item.dispositionUpdatedAt;
        if (item.disposition !== legacyFields.disposition || item.dispositionLocation !== legacyFields.dispositionLocation) {
          dispositionUpdatedAt = new Date().toISOString();
        }

        return {
          ...editingRecord,
          ...legacyFields,
          // Clear irrelevant fields for storage efficiency and data consistency
          protocol: editingRecord.registryType === 'Arrival' ? (editingRecord.customsStatus === 'Cleared' ? 'Cleared Baggage' : 'Non-Cleared / Other') : undefined,
          deliveryAgent: editingRecord.disposition === 'Delivered' ? editingRecord.deliveryAgent : undefined,
          deliveryDateTime: editingRecord.disposition === 'Delivered' ? editingRecord.deliveryDateTime : undefined,
          domesticForwarding: editingRecord.disposition === 'Forwarded' ? editingRecord.domesticForwarding : undefined,
          forwardingFlightNo: (editingRecord.disposition === 'Forwarded' || editingRecord.registryType === 'Departure') ? editingRecord.forwardingFlightNo : undefined,
          forwardingDate: (editingRecord.disposition === 'Forwarded' || editingRecord.registryType === 'Departure') ? editingRecord.forwardingDate : undefined,
          forwardingDestination: (editingRecord.disposition === 'Forwarded' || editingRecord.registryType === 'Departure') ? editingRecord.forwardingDestination : undefined,
          arrivalBelt: editingRecord.disposition === 'Belt 9' ? 'Arrival Belt 9' : undefined,
          handoverAirline: editingRecord.disposition === 'Handover' ? editingRecord.handoverAirline : undefined,
          handoverDateTime: editingRecord.disposition === 'Handover' ? editingRecord.handoverDateTime : undefined,
          warehouseReference: editingRecord.disposition === 'CWC' ? editingRecord.warehouseReference : undefined,
          reexportFlight: editingRecord.disposition === 'Re-export' ? editingRecord.reexportFlight : undefined,
          reexportDestination: editingRecord.disposition === 'Re-export' ? editingRecord.reexportDestination : undefined,
          reexportDate: editingRecord.disposition === 'Re-export' ? editingRecord.reexportDate : undefined,
          preventiveReason: (editingRecord.customsStatus === 'Not Cleared' || editingRecord.customsStatus === 'Marked Preventive') ? editingRecord.preventiveReason : undefined,
          receivedAt,
          dispositionUpdatedAt,
          customsUpdatedAt: item.customsStatus !== legacyFields.customsStatus ? new Date().toISOString() : item.customsUpdatedAt
        } as BaggageRecord;
      }
      return item;
    });

    saveBaggageData(updatedList);
    handleOpenEditDialog(null);
  };

  // Bulk operation actions
  const handleBulkEditSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (selectedIds.length === 0) return;

    const updatedList = baggageList.map(item => {
      if (selectedIds.includes(item.id)) {
        // Redesigned Bulk Update Mapping
        const updatedCustomsStatus = bulkCustomsStatus;
        const updatedDisposition = bulkDisposition;

        let updatedLocation: any = '';
        if (updatedDisposition === 'Delivered') {
          updatedLocation = bulkLocation || 'VVM';
        } else if (updatedDisposition === 'Storage') {
          updatedLocation = 'LHG Office';
        } else if (updatedDisposition === 'Forwarded') {
          const df = bulkDomesticForwarding || 'Air India';
          updatedLocation = df === 'IndiGo' ? 'Indigo' : df === 'SpiceJet' ? 'Spice Jet' : 'Air India';
        } else if (updatedDisposition === 'Belt 9') {
          updatedLocation = 'Belt 9';
        } else if (updatedDisposition === 'Handover') {
          updatedLocation = bulkHandoverAirline || 'Other Airline';
        } else if (updatedDisposition === 'CWC') {
          updatedLocation = 'CWC';
        } else if (updatedDisposition === 'Re-export') {
          updatedLocation = 'Hub Re-export';
        }

        const legacyProtocol = updatedCustomsStatus === 'Cleared' ? 'Cleared Baggage' : 'Non-Cleared / Other';

        return {
          ...item,
          customsStatus: updatedCustomsStatus,
          customsUpdatedAt: new Date().toISOString(),
          disposition: updatedDisposition,
          dispositionLocation: updatedLocation,
          dispositionUpdatedAt: new Date().toISOString(),
          status: 'Received' as const,
          receivedAt: item.receivedAt || new Date().toISOString(),
          // Storage efficiency: clear irrelevant fields
          protocol: (item.registryType === 'Arrival' ? legacyProtocol : undefined) as any,
          preventiveReason: (updatedCustomsStatus === 'Not Cleared' || updatedCustomsStatus === 'Marked Preventive') ? bulkPreventiveReason : undefined,
          deliveryAgent: updatedDisposition === 'Delivered' ? (bulkLocation as any) : undefined,
          deliveryDateTime: updatedDisposition === 'Delivered' ? bulkDeliveryDateTime : undefined,
          domesticForwarding: updatedDisposition === 'Forwarded' ? bulkDomesticForwarding : undefined,
          forwardingFlightNo: (updatedDisposition === 'Forwarded' || item.registryType === 'Departure') ? bulkForwardingFlightNo : undefined,
          forwardingDate: (updatedDisposition === 'Forwarded' || item.registryType === 'Departure') ? bulkForwardingDateNew : undefined,
          forwardingDestination: (updatedDisposition === 'Forwarded' || item.registryType === 'Departure') ? bulkForwardingDestinationNew : undefined,
          arrivalBelt: updatedDisposition === 'Belt 9' ? 'Arrival Belt 9' : undefined,
          handoverAirline: updatedDisposition === 'Handover' ? bulkHandoverAirline : undefined,
          handoverDateTime: updatedDisposition === 'Handover' ? bulkHandoverDateTime : undefined,
          warehouseReference: updatedDisposition === 'CWC' ? bulkWarehouseReference : undefined,
          reexportFlight: updatedDisposition === 'Re-export' ? bulkReexportFlight : undefined,
          reexportDestination: updatedDisposition === 'Re-export' ? bulkReexportDestination : undefined,
          reexportDate: updatedDisposition === 'Re-export' ? bulkReexportDate : undefined,
          storageRemarks: (updatedDisposition === 'Storage' || updatedDisposition === 'CWC') ? bulkStorageRemarks : undefined,
          remarks: bulkRemarks || item.remarks
        } as BaggageRecord;
      }
      return item;
    });

    saveBaggageData(updatedList);
    setSelectedIds([]);
    setShowBulkEdit(false);
  };

  // Bulk Delete - Admin & Operators (Triggers custom modal)
  const handleBulkDelete = () => {
    if (!user) {
      alert('AUTHENTICATION REQUIRED: Please log in to delete records.');
      return;
    }
    setDeleteConfirm({ show: true, type: 'bulk' });
  };

  // Single Item Delete - Admin & Operators (Triggers custom modal)
  const handleDeleteSingle = (id: string) => {
    if (!user) {
      alert('AUTHENTICATION REQUIRED: Please log in to delete records.');
      return;
    }
    setDeleteConfirm({ show: true, type: 'single', id });
  };

  // Executes actual record deletions after modal confirmation (iframe/sandbox-safe)
  const executeDeleteConfirmed = () => {
    if (!user) {
      alert('AUTHENTICATION REQUIRED: Please log in to delete records.');
      setDeleteConfirm({ show: false, type: 'single' });
      return;
    }

    if (deleteConfirm.type === 'bulk') {
      const remaining = baggageList.filter(item => !selectedIds.includes(item.id));
      saveBaggageData(remaining);
      setSelectedIds([]);
    } else if (deleteConfirm.type === 'single' && deleteConfirm.id) {
      const singleId = deleteConfirm.id;
      const remaining = baggageList.filter(item => item.id !== singleId);
      saveBaggageData(remaining);
      // Remove from selection if checked
      setSelectedIds(prev => prev.filter(selectedId => selectedId !== singleId));
    }
    setDeleteConfirm({ show: false, type: 'single' });
  };

  // Toggle multi-select checkbox
  const handleSelectToggle = (id: string) => {
    setSelectedIds(prev => 
      prev.includes(id) ? prev.filter(i => i !== id) : [...prev, id]
    );
  };

  const handleSelectAll = (filteredRecords: BaggageRecord[]) => {
    const filteredIds = filteredRecords.map(r => r.id);
    const allSelected = filteredIds.every(id => selectedIds.includes(id));
    
    if (allSelected) {
      // Unselect only these filtered records
      setSelectedIds(prev => prev.filter(id => !filteredIds.includes(id)));
    } else {
      // Merge selection
      setSelectedIds(prev => {
        const union = new Set([...prev, ...filteredIds]);
        return Array.from(union);
      });
    }
  };

  // Helper to parse pasted text into a grid of string values
  const parsePastedGrid = (text: string): any[][] => {
    if (!text.trim()) return [];
    const lines = text.split('\n');
    return lines.map(line => {
      // Handle both Tab-separated and Comma-separated (CSV style) pasted data
      const isTab = line.includes('\t');
      const cells = isTab 
        ? line.split('\t') 
        : line.split(/,(?=(?:(?:[^"]*"){2})*[^"]*$)/);
      return cells.map(c => c.replace(/^"|"$/g, '').trim());
    });
  };

  // Perform Intelligent Header & Semantic analysis on a grid of rows
  const analyzeGrid = (grid: any[][]) => {
    if (grid.length === 0) return;
    setIsAnalyzing(true);

    try {
      // 1. Detect Header Row: Scan first 5 rows and find highest match rate against mapping dictionary
      let headerRowIdx = 0;
      let maxMatches = -1;

      const scanLimit = Math.min(5, grid.length);
      for (let r = 0; r < scanLimit; r++) {
        let matchCount = 0;
        const row = grid[r];
        if (!row) continue;

        row.forEach(cell => {
          if (!cell) return;
          const cellStr = String(cell).toLowerCase().trim();
          
          // Check if matches any of our dictionary aliases
          const found = mappingDictionary.some(entry => 
            entry.aliases.some(alias => getSimilarityScore(cellStr, alias) > 0.8)
          );
          if (found) matchCount++;
        });

        if (matchCount > maxMatches) {
          maxMatches = matchCount;
          headerRowIdx = r;
        }
      }

      // 2. Extract Headers & Data Rows
      const rawHeaders = grid[headerRowIdx];
      const headers = rawHeaders.map((h, i) => {
        if (h === null || h === undefined || String(h).trim() === '') {
          return `Column ${i + 1}`;
        }
        return String(h).trim();
      });

      const dataRows = grid.slice(headerRowIdx + 1).filter(r => 
        r && r.some(cell => cell !== null && cell !== undefined && String(cell).trim() !== '')
      );

      // 3. For each column, find best mapping field
      const mappings: Record<number, { systemField: string, confidence: number, matchedBy: 'header' | 'semantic' | 'manual' }> = {};
      
      if (isSBHImport) {
        // SBH FIXED PROFILE MAPPING - ONLY 4 FIELDS
        for (let colIdx = 0; colIdx < headers.length; colIdx++) {
          const headerVal = headers[colIdx].trim().toLowerCase();
          let systemField = 'ignore';

          if (headerVal === 'bag tag number') systemField = 'originalTag';
          else if (headerVal === 'pax') systemField = 'name';
          else if (headerVal === 'outbound flight number') systemField = 'flightNo';

          mappings[colIdx] = {
            systemField,
            confidence: systemField !== 'ignore' ? 100 : 0,
            matchedBy: 'header'
          };
        }
      } else {
        for (let colIdx = 0; colIdx < headers.length; colIdx++) {
        const headerVal = headers[colIdx];
        const cleanHeader = headerVal.toLowerCase().replace(/[^a-z0-9]/g, '');

        // Determine if this is a sequence/ID/Serial Number column to prevent incorrect auto-mapping to Weight
        const isSerialOrNoColumn = /^(sno|s\.no|serial|seq|id|index|no\d*|num)$/i.test(headerVal.trim().toLowerCase()) || /serial/i.test(headerVal);

        if (isSerialOrNoColumn) {
          mappings[colIdx] = {
            systemField: 'ignore',
            confidence: 100,
            matchedBy: 'header'
          };
          continue;
        }

        let bestField = 'ignore';
        let bestConfidence = 0;
        let bestMatchedBy: 'header' | 'semantic' | 'manual' = 'header';

        // 3a. Header Aliases Matching
        mappingDictionary.forEach(entry => {
          // If the entry is weight, and it's a serial column, don't map
          if (entry.field === 'weight' && isSerialOrNoColumn) {
            return;
          }
          entry.aliases.forEach(alias => {
            const score = getSimilarityScore(cleanHeader, alias);
            const confPercent = Math.round(score * 100);
            if (confPercent > bestConfidence) {
              bestConfidence = confPercent;
              bestField = entry.field;
              bestMatchedBy = 'header';
            }
          });
        });

        // 3b. Semantic Inspection
        const columnValues = dataRows.slice(0, 15).map(r => r[colIdx] !== undefined ? String(r[colIdx]).trim() : '').filter(Boolean);
        if (columnValues.length > 0) {
          // Calculate semantic matches for major fields
          let matchStats = {
            flightNo: 0,
            originalTag: 0,
            rushTag: 0,
            pir: 0,
            weight: 0,
            damaged: 0,
            ln: 0,
            name: 0,
            destination: 0
          };

          columnValues.forEach(val => {
            const upper = val.toUpperCase();
            
            // Flight regex (e.g. LH760, AI302, EK501)
            if (/^[A-Z]{2,3}\d{3,4}$/i.test(val)) matchStats.flightNo++;
            
            // Original Tag (10 digits)
            if (/^\d{10}$/.test(val)) matchStats.originalTag++;
            
            // Rush Tag (8-12 digits or alphanumeric e.g. LX920394)
            if (/^[A-Z]{2}\d{5,8}$/i.test(val) || (/^[A-Z0-9]+$/i.test(val) && val.length >= 6 && val.length <= 12 && !/^\d+$/.test(val))) matchStats.rushTag++;
            
            // PIR Number (e.g. DELEK12345 or DEL_LX_10294)
            if (/^[A-Z]{3}_?[A-Z]{2}_?\d{4,6}$/i.test(val)) matchStats.pir++;
            
            // Weight (decimal or float)
            if (!isNaN(Number(val)) && Number(val) > 0 && Number(val) < 65) matchStats.weight++;
            
            // Damaged (Y/N/Yes/No)
            if (/^(y|n|yes|no|true|false)$/i.test(val)) matchStats.damaged++;
            
            // Locked (Yes/No/True/False/L01/L02 or free text like PAD, CL, SEAL, combination, padlock, sealed, tsa, zip, cable, tape, wrap, string, none, custom, security strap)
            if (/^(y|n|yes|no|true|false|l\d{2}|pad|cl|seal|padlock|combination lock|sealed|tsa|zip|cable|tape|wrap|string|none|custom|security|locked)$/i.test(val)) matchStats.ln++;

            // Destination (3-letter)
            if (/^[A-Z]{3}$/i.test(val)) matchStats.destination++;

            // Name (alphabetic with spaces and comma, e.g. MUELLER HANS, Amit Kumar)
            if (/^[A-Z\s,-]+$/i.test(val) && val.includes(' ') && val.length > 5) matchStats.name++;
          });

          const totalSamples = columnValues.length;
          
          // Compute semantic scores
          const semanticRates = {
            flightNo: matchStats.flightNo / totalSamples,
            originalTag: matchStats.originalTag / totalSamples,
            rushTag: matchStats.rushTag / totalSamples,
            pir: matchStats.pir / totalSamples,
            weight: matchStats.weight / totalSamples,
            damaged: matchStats.damaged / totalSamples,
            ln: matchStats.ln / totalSamples,
            destination: matchStats.destination / totalSamples,
            name: matchStats.name / totalSamples
          };

          // Compare semantic match to find high-probability fields
          Object.entries(semanticRates).forEach(([field, rate]) => {
            if (field === 'weight' && isSerialOrNoColumn) {
              return;
            }
            if (rate > 0.4) {
              const semConf = Math.round(rate * 85);
              // If semantic confidence is stronger than header confidence, promote it
              if (semConf > bestConfidence) {
                bestConfidence = semConf;
                bestField = field;
                bestMatchedBy = 'semantic';
              } else if (bestField === field && bestConfidence < 95) {
                // Synthesize header + semantic strength
                bestConfidence = Math.min(bestConfidence + 10, 98);
              }
            }
          });
        }

        // Apply strict cut-off for low confidence
        if (bestConfidence < 25) {
          bestField = 'ignore';
          bestConfidence = 0;
        }

        mappings[colIdx] = {
          systemField: bestField,
          confidence: bestConfidence,
          matchedBy: bestMatchedBy
        };
      }
    }

      setExcelHeaders(headers);
      setExcelRows(dataRows);
      setColumnMappings(mappings);
      setImportWizardStep('mapping');
    } catch (e: any) {
      alert(`Analysis failed: ${e.message || e}`);
    } finally {
      setIsAnalyzing(false);
    }
  };

  // Convert raw values and execute import
  const executeIntelligentImport = () => {
    if (excelRows.length === 0) return;
    setImportProgress(10);

    const newImportedRecords: BaggageRecord[] = [];
    let duplicateCount = 0;
    let skippedCount = 0;
    let updatedCount = 0;
    let invalidCount = 0;
    
    // Additional metrics for custom requirements
    let blankWeightCount = 0;
    let invalidWeightCount = 0;
    let recognizedLockCount = 0;
    let unrecognizedLockCount = 0;

    const warningList: string[] = [];

    // Map system fields to Excel indices
    const inverseMappings: Record<string, number> = {};
    Object.entries(columnMappings).forEach(([colIdx, mapping]) => {
      if (mapping.systemField !== 'ignore') {
        inverseMappings[mapping.systemField] = Number(colIdx);
      }
    });

    setImportProgress(40);

    excelRows.forEach((row, rIdx) => {
      const getVal = (field: string, defaultVal: string = ''): string => {
        const idx = inverseMappings[field];
        if (idx !== undefined && row[idx] !== undefined && row[idx] !== null) {
          return String(row[idx]).trim();
        }
        return defaultVal;
      };

      const originalTagVal = getVal('originalTag');
      const rushTagVal = getVal('rushTag');
      const pirVal = getVal('pir');
      const nameVal = getVal('name');
      const flightVal = getVal('flightNo', 'LH760');
      const weightRaw = getVal('weight');
      const damagedRaw = getVal('damaged', 'N');
      const lockedRaw = getVal('ln', '');
      const destVal = getVal('destination', '');
      const sealVal = getVal('seal');
      const remarksVal = getVal('remarks', 'Imported via Intelligent Importer');
      const protocolRaw = getVal('protocol');
      const receivedAtRaw = getVal('receivedAt');

      // Row Validation
      if (!nameVal && !pirVal && !originalTagVal && !rushTagVal) {
        // Skip entirely empty row
        return;
      }

      let hasWarning = false;
      let finalName = nameVal;
      let finalPir = pirVal;

      if (!nameVal) {
        finalName = 'UNKNOWN PASSENGER';
        warningList.push(`Row ${rIdx + 1}: Missing Passenger Name. Auto-filled as 'UNKNOWN PASSENGER'.`);
        hasWarning = true;
      }
      if (!pirVal) {
        finalPir = 'NO PIR';
        warningList.push(`Row ${rIdx + 1}: Missing PIR Reference.`);
        hasWarning = true;
      }

      // Robust Weight Validation Logic
      let parsedWeight: number | undefined = undefined;
      const hasWeightColumn = inverseMappings['weight'] !== undefined;
      if (hasWeightColumn) {
        if (weightRaw !== '') {
          // Strip any trailing unit suffixes like kg, kgs, lbs, lb (case-insensitive) and spaces
          const cleanWeight = weightRaw.replace(/(kg|kgs|lbs|lb)\s*$/i, '').trim();
          const parsed = Number(cleanWeight);
          if (!isNaN(parsed) && parsed >= 0) {
            parsedWeight = parsed;
          } else {
            invalidWeightCount++;
            warningList.push(`Row ${rIdx + 1}: Invalid weight value "${weightRaw}". Left blank.`);
            parsedWeight = undefined;
          }
        } else {
          blankWeightCount++;
          parsedWeight = undefined;
        }
      } else {
        blankWeightCount++;
        parsedWeight = undefined;
      }

      // Convert damaged to 'Y' | 'N'
      let finalDamaged: 'Y' | 'N' = 'N';
      if (/^(y|yes|true|dmg|damaged|1)$/i.test(damagedRaw)) {
        finalDamaged = 'Y';
      }

      // Robust Locked (LN) Expansion Logic
      let finalLocked = '';
      const hasLockColumn = inverseMappings['ln'] !== undefined;
      if (hasLockColumn && lockedRaw !== '') {
        // Find case-insensitive match in lock dictionary
        const keys = Object.keys(lockDictionary);
        const matchedKey = keys.find(k => k.toLowerCase() === lockedRaw.toLowerCase());
        if (matchedKey) {
          finalLocked = lockDictionary[matchedKey];
          recognizedLockCount++;
        } else {
          finalLocked = lockedRaw;
          unrecognizedLockCount++;
        }
      } else {
        finalLocked = lockedRaw; // Preserve original blank if not mentioned
      }

      // Detect Protocol defaults
      let finalProtocol: 'Cleared Baggage' | 'Non-Cleared / Other' = 'Cleared Baggage';
      if (protocolRaw) {
        if (/non|other|uncleared|customs|cwc/i.test(protocolRaw)) {
          finalProtocol = 'Non-Cleared / Other';
        }
      }

      // Validate flight numbers to default options
      const allowedFlights = ['LH760', 'LH762', 'LX146', 'LX2646'];
      let finalFlight = '';
      let sbhReceiptDate: string | undefined = undefined;
      let sbhDestination: string | undefined = undefined;

      if (isSBHImport && flightVal && flightVal.includes('/')) {
        // SBH Format: LH760/06JUL DEL→BOM
        try {
          const parts = flightVal.split('/');
          finalFlight = parts[0].trim().toUpperCase().replace(/\s+/g, '');
          
          if (parts.length > 1) {
            const rest = parts[1].trim();
            const restParts = rest.split(/\s+/);
            const dateStrRaw = restParts[0]; // e.g. "06JUL"
            
            // Extract route (e.g. "DEL→BOM")
            if (restParts.length > 1) {
              const route = restParts[1];
              // Support both → and -> 
              const arrow = route.includes('→') ? '→' : '->';
              if (route.includes(arrow)) {
                sbhDestination = route.split(arrow)[0].trim().toUpperCase();
              }
            }

            // Parse date (e.g. "06JUL")
            // We'll append current year
            const currentYear = new Date().getFullYear();
            const d = new Date(`${dateStrRaw} ${currentYear}`);
            if (!isNaN(d.getTime())) {
              sbhReceiptDate = d.toISOString();
            } else {
              warningList.push(`Row ${rIdx + 1}: Failed to parse SBH date "${dateStrRaw}". Using current date.`);
            }
          }
        } catch (err) {
          warningList.push(`Row ${rIdx + 1}: Error parsing SBH flight field: ${flightVal}. Using defaults.`);
          finalFlight = flightVal.split('/')[0].trim();
        }
      } else if (flightVal) {
        const cleanedFlight = flightVal.toUpperCase().replace(/\s+/g, '');
        const found = allowedFlights.find(f => cleanedFlight.includes(f) || cleanedFlight === f);
        finalFlight = found || flightVal;
      }

      const candidate: BaggageRecord = {
        id: `bag-i-${Date.now()}-${rIdx}-${Math.floor(Math.random() * 1000)}`,
        sno: getVal('sno') || (baggageList.length + newImportedRecords.length + 1).toString(),
        pir: isSBHImport ? '' : finalPir.toUpperCase(),
        name: finalName.toUpperCase(),
        originalTag: originalTagVal,
        rushTag: isSBHImport ? '' : rushTagVal.toUpperCase(),
        flightNo: finalFlight,
        seal: isSBHImport ? '' : sealVal.toUpperCase(),
        ln: isSBHImport ? '' : finalLocked,
        destination: isSBHImport ? (sbhDestination || '') : destVal.toUpperCase(),
        remarks: isSBHImport ? 'SBH Import' : remarksVal,
        status: 'Expected',
        receivedAt: sbhReceiptDate || (receivedAtRaw ? new Date(receivedAtRaw).toISOString() : new Date().toISOString()),
        customsStatus: 'Pending',
        disposition: 'Pending',
        createdAt: new Date().toISOString(),
        weight: isSBHImport ? undefined : parsedWeight,
        damaged: isSBHImport ? undefined : finalDamaged,
        protocol: isSBHImport ? '' : finalProtocol,
        deliveryAgent: (!isSBHImport && finalProtocol === 'Cleared Baggage') ? 'VVM' : undefined,
        storageOption: (!isSBHImport && finalProtocol === 'Cleared Baggage') ? 'Standard Warehousing – LHG Office' : undefined,
        domesticForwarding: (!isSBHImport && finalProtocol === 'Cleared Baggage') ? 'No Forwarding' : undefined,
        arrivalBelt: (!isSBHImport && finalProtocol === 'Non-Cleared / Other') ? 'Arrival Belt 9' : undefined,
        handoverOption: (!isSBHImport && finalProtocol === 'Non-Cleared / Other') ? 'Partner Airlines' : undefined,
        warehouseOption: (!isSBHImport && finalProtocol === 'Non-Cleared / Other') ? 'CWC Warehouse' : undefined,
        reexportOption: (!isSBHImport && finalProtocol === 'Non-Cleared / Other') ? 'Re-export to Carrier Hub' : undefined,
        registryType: activeRegistry === 'Combined' ? 'Arrival' : activeRegistry
      };

      // Check Duplicates against existing local baggage records AND within the current import batch
      const isDuplicateInDb = baggageList.find(item => {
        if (candidate.pir && candidate.pir !== 'NO PIR' && item.pir === candidate.pir) return true;
        if (candidate.originalTag && item.originalTag === candidate.originalTag) return true;
        if (candidate.rushTag && item.rushTag === candidate.rushTag) return true;
        return false;
      });

      const isDuplicateInBatch = newImportedRecords.find(item => {
        if (candidate.pir && candidate.pir !== 'NO PIR' && item.pir === candidate.pir) return true;
        if (candidate.originalTag && item.originalTag === candidate.originalTag) return true;
        if (candidate.rushTag && item.rushTag === candidate.rushTag) return true;
        return false;
      });

      if (isDuplicateInDb || isDuplicateInBatch) {
        duplicateCount++;
        
        // Internal batch duplicates are always skipped (only import first occurrence)
        if (isDuplicateInBatch) {
          skippedCount++;
          warningList.push(`Row ${rIdx + 1}: Duplicate tag "${candidate.originalTag || candidate.rushTag}" found across selected files. Record skipped.`);
          return;
        }

        if (duplicateMode === 'skip') {
          skippedCount++;
          return; // Skip importing this row
        } else if (duplicateMode === 'update' && isDuplicateInDb) {
          // Merge properties and update the existing list
          updatedCount++;
          const idxToUpdate = baggageList.findIndex(x => x.id === isDuplicateInDb.id);
          if (idxToUpdate !== -1) {
            baggageList[idxToUpdate] = {
              ...baggageList[idxToUpdate],
              // Overwrite with incoming non-empty values
              name: candidate.name !== 'UNKNOWN PASSENGER' ? candidate.name : baggageList[idxToUpdate].name,
              flightNo: candidate.flightNo || baggageList[idxToUpdate].flightNo,
              weight: candidate.weight !== undefined ? candidate.weight : baggageList[idxToUpdate].weight,
              damaged: candidate.damaged,
              ln: candidate.ln,
              seal: candidate.seal || baggageList[idxToUpdate].seal,
              remarks: (candidate.remarks || baggageList[idxToUpdate].remarks) + " (Updated via Import)"
            };
          }
          return;
        }
      }

      newImportedRecords.push(candidate);
    });

    setImportProgress(80);

    const mergedList = [...newImportedRecords, ...baggageList];
    saveBaggageData(mergedList);

    setImportProgress(100);
    setImportSummaryResult({
      totalRows: excelRows.length,
      imported: newImportedRecords.length,
      skipped: skippedCount,
      duplicates: duplicateCount,
      invalid: invalidCount,
      blankWeights: blankWeightCount,
      invalidWeights: invalidWeightCount,
      recognizedLocks: recognizedLockCount,
      unrecognizedLocks: unrecognizedLockCount,
      warnings: warningList,
      filesProcessed: isSBHImport ? filesProcessedCount : undefined
    });

    // Clear and step forward
    setImportWizardStep('summary');
    setImportProgress(null);
    setImportProgressMessage('');
  };

  // Paste Text process initiation
  const parsePastedData = () => {
    if (!rawPasteText.trim()) return;
    const grid = parsePastedGrid(rawPasteText);
    analyzeGrid(grid);
  };

  // Uploaded spreadsheet handler (Excel, CSV, TXT)
  const handleCSVUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    setShowImportDialog(true);
    setIsAnalyzing(true);
    setImportProgress(0);
    
    let combinedRows: any[][] = [];
    let firstHeaders: string[] = [];
    let totalRowsProcessed = 0;
    const warningList: string[] = [];

    const fileList = Array.from(files);
    setFilesProcessedCount(fileList.length);
    
    for (let i = 0; i < fileList.length; i++) {
      const file = fileList[i];
      const fileName = file.name.toLowerCase();
      setImportProgress(Math.round(((i) / fileList.length) * 100));
      setImportProgressMessage(`Reading file ${i + 1} of ${fileList.length}: ${file.name}`);

      try {
        const grid = await new Promise<any[][]>((resolve, reject) => {
          const reader = new FileReader();
          reader.onload = (evt) => {
            try {
              const result = evt.target?.result;
              if (!result) {
                resolve([]);
                return;
              }
              const data = new Uint8Array(result as ArrayBuffer);
              const workbook = XLSX.read(data, { type: 'array' });
              const firstSheetName = workbook.SheetNames[0];
              const worksheet = workbook.Sheets[firstSheetName];
              const rows = XLSX.utils.sheet_to_json<any[]>(worksheet, { header: 1 });
              resolve(rows);
            } catch (err) {
              reject(err);
            }
          };
          reader.onerror = reject;
          if (fileName.endsWith('.xls') || fileName.endsWith('.xlsx')) {
            reader.readAsArrayBuffer(file);
          } else {
            // Fallback for non-excel (e.g. CSV)
            const textReader = new FileReader();
            textReader.onload = (evt) => {
              const text = evt.target?.result as string;
              if (!text) { resolve([]); return; }
              const parsed = parsePastedGrid(text);
              resolve(parsed);
            };
            textReader.onerror = reject;
            textReader.readAsText(file);
          }
        });

        if (grid.length > 0) {
          // Detect header for THIS file
          let headerRowIdx = 0;
          let maxMatches = -1;
          const scanLimit = Math.min(5, grid.length);
          for (let r = 0; r < scanLimit; r++) {
            let matchCount = 0;
            const row = grid[r];
            if (!row) continue;
            row.forEach(cell => {
              if (!cell) return;
              const cellStr = String(cell).toLowerCase().trim();
              const found = mappingDictionary.some(entry => 
                entry.aliases.some(alias => alias.toLowerCase() === cellStr) ||
                entry.aliases.some(alias => getSimilarityScore(cellStr, alias) > 0.8)
              );
              if (found) matchCount++;
            });
            if (matchCount > maxMatches) {
              maxMatches = matchCount;
              headerRowIdx = r;
            }
          }

          if (i === 0) {
            firstHeaders = grid[headerRowIdx].map((h, colIdx) => {
              if (h === null || h === undefined || String(h).trim() === '') {
                return `Column ${colIdx + 1}`;
              }
              return String(h).trim();
            });
          }

          const fileDataRows = grid.slice(headerRowIdx + 1).filter(r => 
            r && r.some(cell => cell !== null && cell !== undefined && String(cell).trim() !== '')
          );
          
          combinedRows.push(...fileDataRows);
          totalRowsProcessed += fileDataRows.length;
        }
      } catch (err: any) {
        warningList.push(`Error reading ${file.name}: ${err.message || err}`);
      }
    }

    setImportProgress(100);
    setImportProgressMessage('Analysis complete.');
    
    if (combinedRows.length > 0) {
      // Small delay to let UI update
      setTimeout(() => {
        analyzeGrid([firstHeaders, ...combinedRows]);
      }, 300);
    } else {
      setIsAnalyzing(false);
      setImportProgress(null);
      setImportProgressMessage('');
      alert('No valid data found in selected files.');
    }
  };

  // Altea Import Handler
  const handleAlteaImport = async () => {
    if (!rawAlteaText.trim()) return;

    setShowAlteaPasteWindow(false);
    setIsAlteaImport(true);
    setIsAnalyzing(true);
    setImportProgress(10);
    setImportProgressMessage('Detecting Altea Format...');

    const lines = rawAlteaText.split(/\r?\n/);
    const firstLine = lines.find(l => l.trim()) || '';
    
    let format: 'cryptic' | 'content' | null = null;
    if (rawAlteaText.includes('JFE Screen Copy - Plain Text') || rawAlteaText.includes('LOADING STATUS')) {
      format = 'cryptic';
    } else if (rawAlteaText.includes('LIST OF:') || rawAlteaText.includes('BAG LIST')) {
      format = 'content';
    }

    if (!format) {
      setIsAnalyzing(false);
      setImportProgress(null);
      alert('Unrecognized Altea format. Please ensure you copied the Cryptic or Content output correctly.');
      return;
    }

    setImportProgressMessage(`Parsing Altea ${format} data...`);
    
    const newImportedRecords: BaggageRecord[] = [];
    let flightNo = '';
    let receiptDateRaw = '';
    const warningList: string[] = [];

    // 1. Header parsing (Flight & Date)
    for (let i = 0; i < Math.min(30, lines.length); i++) {
      const line = lines[i].trim();
      // Look for flight LH760 or LX146 etc.
      const flightMatch = line.match(/\b([A-Z]{2,3}\d{3,4})\b/i);
      if (flightMatch && !flightNo) {
        flightNo = flightMatch[1].toUpperCase();
      }
      // Look for 05JUL
      const dateMatch = line.match(/\b(\d{2}[A-Z]{3})\b/i);
      if (dateMatch && !receiptDateRaw) {
        receiptDateRaw = dateMatch[1].toUpperCase();
      }
    }

    // 2. Date Math: Input 05JUL -> Store 06JUL
    let finalReceiptDateISO = new Date().toISOString();
    if (receiptDateRaw) {
      try {
        const currentYear = new Date().getFullYear();
        const d = new Date(`${receiptDateRaw} ${currentYear}`);
        if (!isNaN(d.getTime())) {
          // Add one calendar day as per requirement
          d.setDate(d.getDate() + 1);
          finalReceiptDateISO = d.toISOString();
        }
      } catch (e) {
        warningList.push(`Failed to calculate +1 day for Altea date: ${receiptDateRaw}`);
      }
    }

    // 3. Record Extraction
    if (format === 'cryptic') {
      // Cryptic Output parsing
      // Pattern: 0014-AC-246669
      lines.forEach((line, idx) => {
        const tagMatch = line.match(/(\d{4})-([A-Z]{2})-(\d{6})/);
        if (tagMatch) {
          const originalTag = tagMatch[2] + tagMatch[3];
          let rushTag = '';
          let passengerName = '';
          
          // Check for Loading Status N
          if (line.includes('Loading Status N') || line.match(/Status\s+N/i)) {
            rushTag = 'Not Loaded';
          }

          // Extract name from cryptic pipe-delimited format (usually the 14th column)
          if (line.includes('|')) {
            const pipeParts = line.split('|');
            if (pipeParts.length > 14) {
              passengerName = pipeParts[14].trim().toUpperCase();
            }
          }

          newImportedRecords.push({
            id: `alt-c-${Date.now()}-${idx}`,
            sno: (baggageList.length + newImportedRecords.length + 1).toString(),
            pir: '',
            name: passengerName || '',
            originalTag,
            rushTag,
            flightNo: flightNo || 'LH760',
            seal: '',
            ln: '',
            destination: '',
            remarks: 'Altea Import',
            status: 'Expected',
            receivedAt: finalReceiptDateISO,
            customsStatus: 'Pending',
            disposition: 'Pending',
            createdAt: new Date().toISOString(),
            protocol: '',
            registryType: 'Arrival'
          });
        }
      });
    } else {
      // Content Output parsing
      // Pattern: 2LH 803033
      lines.forEach((line, idx) => {
        const tagMatch = line.match(/(\d)([A-Z]{2})\s*(\d{6})/);
        if (tagMatch) {
          const originalTag = tagMatch[2] + tagMatch[3];
          let passengerName = '';
          
          // Improved Name extraction for Content Output
          // Example: 2LH 803033  DEL  A   CAVANAUGH
          // The name usually follows the destination and status columns
          const lineParts = line.trim().split(/\s+/);
          if (lineParts.length >= 5) {
            // Index 0,1 is the tag (2LH 803033), 2 is DEST, 3 is STS, 4+ is NAME
            passengerName = lineParts.slice(4).join(' ').trim().toUpperCase();
            // Clean up any trailing status info if present
            passengerName = passengerName.split(/\s{2,}/)[0].trim();
          }

          newImportedRecords.push({
            id: `alt-o-${Date.now()}-${idx}`,
            sno: (baggageList.length + newImportedRecords.length + 1).toString(),
            pir: '',
            name: passengerName || '',
            originalTag,
            rushTag: '',
            flightNo: flightNo || 'LH760',
            seal: '',
            ln: '',
            destination: '',
            remarks: 'Altea Import',
            status: 'Expected',
            receivedAt: finalReceiptDateISO,
            customsStatus: 'Pending',
            disposition: 'Pending',
            createdAt: new Date().toISOString(),
            protocol: '',
            registryType: 'Arrival'
          });
        }
      });
    }

    if (newImportedRecords.length === 0) {
      setIsAnalyzing(false);
      setImportProgress(null);
      alert('No baggage records could be extracted from the pasted text. Please verify the Altea output format.');
      return;
    }

    // Show Preview instead of immediate commit
    setAlteaPreviewRecords(newImportedRecords);
    // Select all by default
    const allIds = new Set(newImportedRecords.map(r => r.id));
    setSelectedAlteaRowIds(allIds);
    
    setIsAnalyzing(false);
    setImportProgress(null);
    setImportProgressMessage('');
    setShowAlteaPreview(true);
  };

  const handleCameraCapture = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsOcrLoading(true);
    setOcrError(null);
    setImportProgress(30);
    setImportProgressMessage('Scanning document with Gemini Vision...');

    try {
      const reader = new FileReader();
      reader.readAsDataURL(file);
      reader.onload = async () => {
        try {
          const base64Image = reader.result as string;
          
          const response = await fetch('/api/ocr', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ image: base64Image, mode: ocrTargetMode }),
          });

          const result = await response.json();

          if (result.error) {
            throw new Error(result.error);
          }

          if (ocrTargetMode === 'altea') {
            setRawAlteaText(result.text || '');
            setShowAlteaPasteWindow(true);
            setImportProgress(null);
            setImportProgressMessage('');
            setIsOcrLoading(false);
          } else {
            // Handle table data (BDO/SBH)
            if (result.data) {
              const headers = ['Flight No', 'Date', 'Original Tag', 'Passenger Name', 'Weight', 'Rush Status', 'Remarks'];
              const rows = result.data.map((obj: any) => [
                obj.flightNo || '',
                obj.receivedAt || '',
                obj.originalTag || '',
                obj.name || '',
                obj.weight || '',
                obj.rushTag || '',
                obj.remarks || ''
              ]);
              
              setImportProgress(60);
              setImportProgressMessage('Mapping OCR fields to database...');
              
              setTimeout(() => {
                analyzeGrid([headers, ...rows]);
                setImportProgress(null);
                setImportProgressMessage('');
                setIsOcrLoading(false);
              }, 800);
            } else {
              throw new Error("Could not extract table data from image.");
            }
          }
        } catch (err: any) {
          setOcrError(err.message || 'OCR processing failed');
          setIsOcrLoading(false);
          setImportProgress(null);
        }
      };
    } catch (err: any) {
      setOcrError(err.message || 'Image reading failed');
      setIsOcrLoading(false);
      setImportProgress(null);
    } finally {
      if (cameraInputRef.current) cameraInputRef.current.value = '';
    }
  };

  const handleCommitAlteaImport = async () => {
    setIsAnalyzing(true);
    setImportProgress(20);
    setImportProgressMessage('Committing selected Altea records...');

    // Filter only selected AND valid records
    const selectedRecords = alteaPreviewRecords.filter(r => 
      selectedAlteaRowIds.has(r.id) && 
      r.originalTag && 
      r.originalTag.length >= 6
    );
    
    const invalidCount = selectedAlteaRowIds.size - selectedRecords.length;
    
    let duplicateCount = 0;
    let skippedCount = 0;
    let updatedCount = 0;
    const finalImportBatch: BaggageRecord[] = [];
    const localBaggageList = [...baggageList];

    selectedRecords.forEach((candidate) => {
      const isDuplicateInDb = localBaggageList.find(item => 
        candidate.originalTag && item.originalTag === candidate.originalTag
      );

      const isDuplicateInBatch = finalImportBatch.find(item => 
        candidate.originalTag && item.originalTag === candidate.originalTag
      );

      if (isDuplicateInDb || isDuplicateInBatch) {
        duplicateCount++;
        if (isDuplicateInBatch) {
          skippedCount++;
          return;
        }

        if (duplicateMode === 'skip') {
          skippedCount++;
          return;
        } else if (duplicateMode === 'update' && isDuplicateInDb) {
          updatedCount++;
          const idxToUpdate = localBaggageList.findIndex(x => x.id === isDuplicateInDb.id);
          if (idxToUpdate !== -1) {
            localBaggageList[idxToUpdate] = {
              ...localBaggageList[idxToUpdate],
              name: candidate.name || localBaggageList[idxToUpdate].name,
              flightNo: candidate.flightNo || localBaggageList[idxToUpdate].flightNo,
              rushTag: candidate.rushTag || localBaggageList[idxToUpdate].rushTag,
              remarks: localBaggageList[idxToUpdate].remarks + " (Updated via Altea Import)"
            };
          }
          return;
        }
      }
      finalImportBatch.push(candidate);
    });

    const mergedList = [...finalImportBatch, ...localBaggageList];
    saveBaggageData(mergedList);

    setImportProgress(100);
    setImportSummaryResult({
      totalRows: alteaPreviewRecords.length,
      imported: finalImportBatch.length,
      skipped: skippedCount,
      duplicates: duplicateCount,
      invalid: invalidCount,
      blankWeights: finalImportBatch.length,
      invalidWeights: 0,
      recognizedLocks: 0,
      unrecognizedLocks: 0,
      warnings: [],
      filesProcessed: 1
    });

    setImportWizardStep('summary');
    setShowImportDialog(true);
    setShowAlteaPreview(false);
    setIsAnalyzing(false);
    setImportProgress(null);
    setImportProgressMessage('');
    setRawAlteaText('');
  };

  // Reset all filters to default
  const handleClearFilters = () => {
    setActiveFilter({ type: 'all' });
    setSearchTerm('');
    setFilterFlights([]);
    setFilterRemarks('');
    setFilterDateRange({ type: 'all' });
    setFilterRegistry('Combined');
  };

  // Filter records based on current search term, flight select, and active bento dashboard button
  const filteredBaggage = useMemo(() => {
    return baggageList.filter(item => {
      // 1. Text Search matches Name, PIR, tags, remarks, destination
      const searchLower = searchTerm.toLowerCase();
      const matchesSearch = !searchTerm || 
        item.name.toLowerCase().includes(searchLower) ||
        item.pir.toLowerCase().includes(searchLower) ||
        item.originalTag.toLowerCase().includes(searchLower) ||
        item.rushTag.toLowerCase().includes(searchLower) ||
        item.remarks.toLowerCase().includes(searchLower) ||
        item.destination.toLowerCase().includes(searchLower);

      // 2. Flight select matches (Enhanced)
      const matchesFlight = filterFlights.length === 0 || filterFlights.some(f => 
        item.flightNo.toLowerCase().includes(f.toLowerCase())
      );

      // 3. Registry Filter matches
      const matchesRegistry = filterRegistry === 'Combined' || item.registryType === filterRegistry;

      // 4. Date Range matches (Enhanced)
      let matchesDate = true;
      if (filterDateRange.type !== 'all' && item.receivedAt) {
        const receiptDate = new Date(item.receivedAt);
        const today = new Date();
        today.setHours(0, 0, 0, 0);
        
        const itemDay = new Date(receiptDate);
        itemDay.setHours(0, 0, 0, 0);

        if (filterDateRange.type === 'today') {
          matchesDate = itemDay.getTime() === today.getTime();
        } else if (filterDateRange.type === 'yesterday') {
          const yesterday = new Date(today);
          yesterday.setDate(yesterday.getDate() - 1);
          matchesDate = itemDay.getTime() === yesterday.getTime();
        } else if (filterDateRange.type === 'last7') {
          const last7 = new Date(today);
          last7.setDate(last7.getDate() - 7);
          matchesDate = itemDay >= last7;
        } else if (filterDateRange.type === 'last30') {
          const last30 = new Date(today);
          last30.setDate(last30.getDate() - 30);
          matchesDate = itemDay >= last30;
        } else if (filterDateRange.type === 'thisMonth') {
          matchesDate = receiptDate.getMonth() === today.getMonth() && receiptDate.getFullYear() === today.getFullYear();
        } else if (filterDateRange.type === 'prevMonth') {
          const prevMonthDate = new Date(today);
          prevMonthDate.setMonth(prevMonthDate.getMonth() - 1);
          matchesDate = receiptDate.getMonth() === prevMonthDate.getMonth() && receiptDate.getFullYear() === prevMonthDate.getFullYear();
        } else if (filterDateRange.type === 'custom' && filterDateRange.from && filterDateRange.to) {
          const from = new Date(filterDateRange.from);
          from.setHours(0, 0, 0, 0);
          const to = new Date(filterDateRange.to);
          to.setHours(23, 59, 59, 999);
          matchesDate = receiptDate >= from && receiptDate <= to;
        }
      }

      // 5. Bento Dashboard buttons filter (Existing)
      let matchesDashboard = true;
      if (activeFilter.type === 'expected') {
        matchesDashboard = item.status === 'Expected';
      } else if (activeFilter.type === 'arrived') {
        matchesDashboard = item.status === 'Received';
      } else if (activeFilter.type === 'non-arrival') {
        matchesDashboard = item.status === 'Expected';
      } else if (activeFilter.type === 'cleared') {
        matchesDashboard = item.status === 'Received' && item.customsStatus === 'Cleared';
      } else if (activeFilter.type === 'not-cleared') {
        matchesDashboard = item.status === 'Received' && (item.customsStatus === 'Not Cleared' || item.customsStatus === 'Marked Preventive');
      } else if (activeFilter.type === 'location') {
        matchesDashboard = item.status === 'Received' && item.dispositionLocation === activeFilter.value;
      } else if (activeFilter.type === 'alerts') {
        matchesDashboard = item.status === 'Received' && getDaysInStorage(item) >= 3;
      }

      // 6. Remarks Filter
      const matchesRemarks = !filterRemarks || (item.remarks && item.remarks.toLowerCase().includes(filterRemarks.toLowerCase()));

      return matchesSearch && matchesFlight && matchesRegistry && matchesDate && matchesDashboard && matchesRemarks;
    });
  }, [baggageList, searchTerm, filterFlights, filterRemarks, filterDateRange, filterRegistry, activeFilter, getDaysInStorage]);

  // Alert list computed property
  const activeAlerts = useMemo(() => {
    const list3Day: BaggageRecord[] = [];
    const list5Day: BaggageRecord[] = [];
    
    baggageList.forEach(item => {
      const days = getDaysInStorage(item);
      if (days >= 5) {
        list5Day.push(item);
      } else if (days >= 3) {
        list3Day.push(item);
      }
    });

    return {
      list3Day,
      list5Day,
      totalAlerts: list3Day.length + list5Day.length
    };
  }, [baggageList, getDaysInStorage]);

  // Results Summary
  const summary = useMemo(() => {
    const totalRecords = baggageList.length;
    const filteredRecords = filteredBaggage.length;
    const arrival = filteredBaggage.filter(b => b.registryType === 'Arrival').length;
    const departure = filteredBaggage.filter(b => b.registryType === 'Departure').length;
    return { totalRecords, filteredRecords, arrival, departure };
  }, [baggageList, filteredBaggage]);

  // Export to Excel
  const generateExportWorkbook = () => {
    const worksheet = XLSX.utils.json_to_sheet(filteredBaggage.map(r => ({
      'Registry Type': r.registryType,
      'Flight No': r.flightNo,
      'Date of Receipt': r.receivedAt ? new Date(r.receivedAt).toLocaleString('en-GB', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '',
      'Date Arrived': r.createdAt,
      'Original Tag': r.originalTag,
      'Rush Tag': r.rushTag,
      'Passenger Name': r.name,
      'PIR': r.pir,
      'Weight': r.weight,
      'Locked': r.ln,
      'Damaged': r.damaged,
      'Disposition': r.disposition,
      'Storage Location': r.dispositionLocation,
      'Status': r.status,
      'Remarks': r.remarks,
    })));
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Baggage Report');
    return workbook;
  };

  const exportToExcel = () => {
    const workbook = generateExportWorkbook();
    XLSX.writeFile(workbook, `Baggage_Report_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  const handleShareViaWhatsApp = async () => {
    const fileName = `Baggage_Report_${new Date().toISOString().split('T')[0]}.xlsx`;
    const workbook = generateExportWorkbook();
    const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const file = new File([blob], fileName, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

    if (navigator.share && navigator.canShare && navigator.canShare({ files: [file] })) {
      try {
        await navigator.share({
          files: [file],
          title: 'Baggage Report',
          text: 'Please find the attached Baggage Report.'
        });
      } catch (err) {
        console.error('Share failed:', err);
        window.open(`https://wa.me/?text=${encodeURIComponent('Baggage Report ' + fileName)}`, '_blank');
      }
    } else {
      window.open(`https://wa.me/?text=${encodeURIComponent('Baggage Report ' + fileName)}`, '_blank');
      exportToExcel();
    }
  };

  const handleShareViaEmail = async () => {
    const subject = 'Left-Behind Baggage Report';
    const body = 'Please find the attached baggage report.';
    window.location.href = `mailto:?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
    exportToExcel();
  };

  const copyToClipboardTable = async () => {
    const headers = ['Flight No', 'Date', 'Original Tag', 'Name', 'Status', 'Disposition'];
    let text = headers.join('\t') + '\n';
    filteredBaggage.forEach(r => {
      text += [
        r.flightNo,
        r.receivedAt ? new Date(r.receivedAt).toLocaleDateString() : '',
        r.originalTag,
        r.name,
        r.status,
        r.disposition
      ].join('\t') + '\n';
    });

    try {
      await navigator.clipboard.writeText(text);
      alert('Report data copied to clipboard as table!');
    } catch (err) {
      console.error('Failed to copy:', err);
      alert('Failed to copy to clipboard.');
    }
  };

  // Force trigger demo backdate
  const forceBackdateForAlert = (id: string, daysAgo: number) => {
    const updated = baggageList.map(item => {
      if (item.id === id) {
        const backdatedDate = new Date(Date.now() - daysAgo * 24 * 60 * 60 * 1000).toISOString();
        return {
          ...item,
          status: 'Received' as const,
          receivedAt: backdatedDate,
          disposition: 'Storage' as const,
          dispositionLocation: item.dispositionLocation || 'LHG Office',
          dispositionUpdatedAt: backdatedDate
        };
      }
      return item;
    });
    saveBaggageData(updated);
    alert(`Success: Backdated baggage arrival to ${daysAgo} days ago to simulate storage alerting thresholds!`);
  };

  // Ref for the Menu toggle button to return focus when closed via keyboard/ESC
  const menuButtonRef = useRef<HTMLButtonElement>(null);

  // Setup navigation menu structure
  const menuItems = [
    {
      id: 'dashboard' as const,
      label: 'Live Dashboard',
      icon: LayoutDashboard,
      submenu: [
        { label: 'All Overview', filterType: 'all' },
        { label: 'Arrived & Received', filterType: 'arrived' },
        { label: 'Non-Arrivals', filterType: 'non-arrival' },
        { label: 'Customs Cleared', filterType: 'cleared' },
        { label: 'Customs Held', filterType: 'not-cleared' }
      ]
    },
    {
      id: 'scanner' as const,
      label: 'Gate Scanner & Tools',
      icon: QrCode
    },
    {
      id: 'protocol' as const,
      label: 'Operations Protocol',
      icon: Briefcase
    },
    {
      id: 'registry' as const,
      label: 'Operations Registry',
      icon: Database,
      submenu: [
        { label: 'LHG Office Storage', filterType: 'location', filterValue: 'LHG Office' },
        { label: 'BMA Stalls', filterType: 'location', filterValue: 'BMA' },
        { label: 'Level 4 Checks', filterType: 'location', filterValue: 'Level 4 Checks' },
        { label: 'CWC Warehouse', filterType: 'location', filterValue: 'CWC' },
        { label: 'Belt 9 Storage', filterType: 'location', filterValue: 'Belt 9' }
      ]
    }
  ];

  // Helper to change sections
  const handleSectionSelect = (sectionId: 'dashboard' | 'scanner' | 'protocol' | 'registry', isMobileClick = false) => {
    setActiveSection(sectionId);
    if (isMobileClick) {
      setIsMenuOpen(false);
    }
  };

  // Keyboard accessibility: ESC key closes menu, focus returns to Menu button
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        setIsMenuOpen(false);
        menuButtonRef.current?.focus();
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, []);

  // Render empty container on server side to ensure zero hydration mismatches from client-injected iframe attributes or local storage states
  if (!mounted) {
    return <div className="min-h-screen bg-slate-950" />;
  }

  // If user is not logged in, render the custom modern Login screen
  if (!user) {
    return (
      <div id="login-container" className="min-h-screen flex items-center justify-center bg-slate-950 font-sans p-6 relative overflow-hidden">
        {/* Background visual graphics */}
        <div className="absolute top-[-20%] left-[-10%] w-[50%] h-[50%] rounded-full bg-blue-900/10 blur-[120px]" />
        <div className="absolute bottom-[-20%] right-[-10%] w-[50%] h-[50%] rounded-full bg-indigo-900/10 blur-[120px]" />

        <div id="login-card" className="w-full max-w-md bg-slate-900/80 border border-slate-800 backdrop-blur-md rounded-2xl shadow-2xl p-8 relative z-10">
          <div className="text-center mb-8">
            <div className="mx-auto w-14 h-14 bg-gradient-to-tr from-blue-600 to-indigo-500 rounded-xl flex items-center justify-center text-white font-bold text-2xl shadow-lg shadow-indigo-500/20 mb-4">
              <Briefcase className="w-7 h-7" />
            </div>
            <h1 className="text-2xl font-bold text-slate-100 tracking-tight">Rush Baggage Wizard</h1>
            <p className="text-slate-400 text-sm mt-1">Airline Operational Baggage & Customs Registry</p>
          </div>

          <form onSubmit={handleLogin} className="space-y-6">
            {loginError && (
              <div id="login-error" className="bg-red-500/10 border border-red-500/20 text-red-400 text-sm p-3 rounded-lg flex items-center gap-2">
                <AlertCircle className="w-4 h-4 shrink-0" />
                <span>{loginError}</span>
              </div>
            )}

            <div className="space-y-2">
              <label className="text-xs font-semibold text-slate-300 uppercase tracking-wider block">Username</label>
              <div className="relative">
                <User className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-500" />
                <input
                  type="text"
                  required
                  value={usernameInput}
                  onChange={(e) => setUsernameInput(e.target.value)}
                  placeholder="lh or admin"
                  className="w-full bg-slate-950 border border-slate-800 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 text-slate-200 pl-10 pr-4 py-2.5 rounded-lg text-sm transition"
                />
              </div>
            </div>

            <div className="space-y-2">
              <label className="text-xs font-semibold text-slate-300 uppercase tracking-wider block">Password</label>
              <div className="relative">
                <Lock className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-500" />
                <input
                  type="password"
                  required
                  value={passwordInput}
                  onChange={(e) => setPasswordInput(e.target.value)}
                  placeholder="Enter password"
                  className="w-full bg-slate-950 border border-slate-800 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 text-slate-200 pl-10 pr-4 py-2.5 rounded-lg text-sm transition"
                />
              </div>
            </div>

            <button
              type="submit"
              className="w-full bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-500 hover:to-indigo-500 text-white font-medium py-3 rounded-lg text-sm transition shadow-lg shadow-indigo-600/15"
            >
              Secure Operational Login
            </button>
          </form>

          {/* Removed default operator accounts display */}
        </div>
      </div>
    );
  }

  // Once logged in, show the operational dashboard and workspace
  return (
    <div className="min-h-screen bg-slate-950 text-slate-100 font-sans flex flex-col">
      {/* 1. Sticky Header Navigation Bar */}
      <header className="bg-slate-900 border-b border-slate-800 sticky top-0 z-40 px-4 py-3 shadow-md">
        <div className="max-w-7xl mx-auto flex justify-between items-center gap-4">
          <div className="flex items-center gap-3">
            {/* ☰ Menu Button */}
            <button
              ref={menuButtonRef}
              onClick={() => setIsMenuOpen(!isMenuOpen)}
              className="p-2 -ml-2 text-slate-300 hover:text-white hover:bg-slate-800 focus:outline-none focus:ring-2 focus:ring-indigo-500 rounded-lg flex items-center gap-1.5 transition active:scale-95 cursor-pointer"
              aria-label="Toggle Navigation Menu"
              aria-expanded={isMenuOpen}
            >
              <Menu className="w-6 h-6" />
              <span className="text-xs font-bold uppercase tracking-wider hidden sm:inline">Menu</span>
            </button>

            <div className="flex items-center gap-2">
              <div className="w-8 h-8 bg-gradient-to-tr from-blue-600 to-indigo-500 rounded-lg flex items-center justify-center text-white font-semibold shadow">
                <Briefcase className="w-4 h-4" />
              </div>
              <div>
                <h1 className="text-sm font-bold text-slate-100 tracking-tight leading-none">Rush Baggage Wizard</h1>
                <p className="text-[10px] text-slate-400 mt-0.5 hidden xs:block">Swiss International Air Lines Ground Operations & Customs Registry</p>
              </div>
            </div>
          </div>

          <div className="flex items-center gap-3">
            <div className="flex items-center gap-2 px-2.5 py-1 bg-slate-950/80 rounded-lg border border-slate-800 text-[10px]">
              <span className={`w-2 h-2 rounded-full ${user === 'admin' ? 'bg-red-500 animate-pulse' : 'bg-emerald-500'}`} />
              <span className="text-slate-400">Role: </span>
              <span className="font-bold text-slate-100 uppercase font-mono">{user}</span>
            </div>

            <button
              onClick={handleLogout}
              className="flex items-center gap-1 px-2.5 py-1 bg-slate-850 hover:bg-slate-800 text-slate-300 rounded-lg text-[10px] font-medium transition cursor-pointer"
            >
              <LogOut className="w-3 h-3" />
              <span>Logout</span>
            </button>
          </div>
        </div>
      </header>

      {/* Main Workspace Frame with Sidebar + Main Content */}
      <div className="flex flex-1 relative overflow-hidden">
        {/* DESKTOP COLLAPSIBLE SIDEBAR */}
        <aside 
          className={`hidden md:flex flex-col bg-slate-900 border-r border-slate-800 transition-all duration-300 shrink-0 ${
            isMenuOpen ? 'w-64' : 'w-16'
          }`}
        >
          <div className="flex-1 py-6 flex flex-col justify-between">
            <nav className="space-y-1.5 px-3">
              {menuItems.map(item => {
                const isActive = activeSection === item.id;
                return (
                  <div key={item.id} className="space-y-1">
                    <button
                      onClick={() => handleSectionSelect(item.id)}
                      className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-xs font-semibold transition-all cursor-pointer ${
                        isActive 
                          ? 'bg-indigo-600 text-white shadow shadow-indigo-600/20' 
                          : 'text-slate-400 hover:text-slate-200 hover:bg-slate-800/50'
                      }`}
                      title={item.label}
                    >
                      <item.icon className={`w-4 h-4 shrink-0 ${isActive ? 'text-white' : 'text-slate-400'}`} />
                      {isMenuOpen && <span className="truncate">{item.label}</span>}
                    </button>

                    {/* Submenu rendering */}
                    {isMenuOpen && item.submenu && (
                      <div className="pl-7 pr-2 py-1 space-y-1 border-l border-slate-800/50 ml-5">
                        {item.submenu.map(sub => {
                          const isSubActive = activeSection === item.id && activeFilter.type === sub.filterType && ((sub as any).filterValue === undefined || activeFilter.value === (sub as any).filterValue);
                          return (
                            <button
                              key={sub.label}
                              onClick={() => {
                                handleSectionSelect(item.id);
                                if (sub.filterType === 'location') {
                                  setActiveFilter({ type: 'location', value: (sub as any).filterValue });
                                } else {
                                  setActiveFilter({ type: sub.filterType as any });
                                }
                              }}
                              className={`w-full text-left px-2 py-1.5 rounded text-[11px] font-medium transition cursor-pointer ${
                                isSubActive
                                  ? 'text-indigo-400 bg-indigo-950/20 font-bold'
                                  : 'text-slate-500 hover:text-slate-300'
                              }`}
                            >
                              {sub.label}
                            </button>
                          );
                        })}
                      </div>
                    )}
                  </div>
                );
              })}
            </nav>
            <div className="px-3 border-t border-slate-800 pt-4">
              <div className={`text-[10px] text-slate-500 font-mono ${isMenuOpen ? 'block' : 'hidden'}`}>
                SYSTEM RECON ACTIVE
              </div>
            </div>
          </div>
        </aside>

        {/* MOBILE SLIDE-OUT DRAWER */}
        {isMenuOpen && (
          <div 
            className="md:hidden fixed inset-0 bg-black/60 backdrop-blur-xs z-30 transition-opacity"
            onClick={() => setIsMenuOpen(false)}
          />
        )}
        <aside 
          className={`md:hidden fixed inset-y-0 left-0 z-40 w-72 bg-slate-900 border-r border-slate-800 flex flex-col justify-between py-6 transition-transform duration-300 transform shadow-2xl ${
            isMenuOpen ? 'translate-x-0' : '-translate-x-full'
          }`}
        >
          <div className="space-y-6">
            <div className="px-6 flex justify-between items-center">
              <div className="flex items-center gap-2">
                <div className="w-8 h-8 bg-gradient-to-tr from-blue-600 to-indigo-500 rounded-lg flex items-center justify-center text-white">
                  <Briefcase className="w-4 h-4" />
                </div>
                <div>
                  <h2 className="text-xs font-bold uppercase tracking-wider text-slate-100">Baggage Menu</h2>
                  <p className="text-[10px] text-slate-500">Swiss Ground Operations</p>
                </div>
              </div>
              <button 
                onClick={() => {
                  setIsMenuOpen(false);
                  menuButtonRef.current?.focus();
                }}
                className="p-1.5 hover:bg-slate-800 rounded-lg text-slate-400 hover:text-slate-200 cursor-pointer"
                aria-label="Close menu"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <nav className="space-y-1.5 px-4">
              {menuItems.map(item => {
                const isActive = activeSection === item.id;
                return (
                  <div key={item.id} className="space-y-1">
                    <button
                      onClick={() => handleSectionSelect(item.id, true)}
                      className={`w-full flex items-center gap-3 px-3 py-3 rounded-lg text-xs font-semibold transition-all cursor-pointer ${
                        isActive 
                          ? 'bg-indigo-600 text-white shadow shadow-indigo-600/20' 
                          : 'text-slate-400 hover:text-slate-200 hover:bg-slate-800/50'
                      }`}
                    >
                      <item.icon className="w-4 h-4 shrink-0" />
                      <span>{item.label}</span>
                    </button>

                    {item.submenu && (
                      <div className="pl-7 pr-2 py-1 space-y-1.5 border-l border-slate-800/50 ml-5">
                        {item.submenu.map(sub => {
                          const isSubActive = activeSection === item.id && activeFilter.type === sub.filterType && ((sub as any).filterValue === undefined || activeFilter.value === (sub as any).filterValue);
                          return (
                            <button
                              key={sub.label}
                              onClick={() => {
                                handleSectionSelect(item.id, true);
                                if (sub.filterType === 'location') {
                                  setActiveFilter({ type: 'location', value: (sub as any).filterValue });
                                } else {
                                  setActiveFilter({ type: sub.filterType as any });
                                }
                              }}
                              className={`w-full text-left px-2 py-2 rounded text-xs font-medium transition cursor-pointer ${
                                isSubActive
                                  ? 'text-indigo-400 bg-indigo-950/20 font-bold'
                                  : 'text-slate-500 hover:text-slate-300'
                              }`}
                            >
                              {sub.label}
                            </button>
                          );
                        })}
                      </div>
                    )}
                  </div>
                );
              })}
            </nav>
          </div>
          <div className="px-6 border-t border-slate-800 pt-4">
            <span className="text-[10px] text-slate-500 font-mono uppercase tracking-wider block">
              Swiss Air Operations v1.0
            </span>
          </div>
        </aside>

        {/* INDEPENDENTLY SCROLLABLE CONTENT WORKSPACE */}
        <main className="flex-1 overflow-y-auto px-4 sm:px-6 py-6 md:py-8 space-y-6 w-full">
        
        {/* 2. Urgent Reminders Banner Section */}
        {activeAlerts.totalAlerts > 0 && (
          <div id="storage-alerts-banner" className="bg-gradient-to-r from-amber-950/40 to-red-950/40 border border-amber-800/40 rounded-xl p-4 flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 bg-amber-500/20 text-amber-500 rounded-lg flex items-center justify-center">
                <AlertTriangle className="w-5 h-5" />
              </div>
              <div>
                <h3 className="font-bold text-amber-200 text-sm">Critical Storage Threshold Alerts Detected</h3>
                <p className="text-xs text-slate-300 mt-0.5">
                  We have <span className="font-bold text-amber-300">{activeAlerts.list3Day.length}</span> bags in storage for 3+ days, and <span className="font-bold text-red-400 underline">{activeAlerts.list5Day.length}</span> urgent bags at 5+ days requiring immediate re-export.
                </p>
              </div>
            </div>
            <div className="flex gap-2 shrink-0">
              <button
                onClick={() => {
                  setActiveSection('registry');
                  setActiveFilter({ type: 'alerts' });
                }}
                className="px-3 py-1.5 bg-amber-500/20 hover:bg-amber-500/30 text-amber-300 rounded-lg text-xs font-semibold border border-amber-500/30 transition cursor-pointer"
              >
                View Active Alerts
              </button>
            </div>
          </div>
        )}

        {/* 3. Auto-Purge Notification Area */}
        {purgeAlertCount > 0 && (
          <div className="bg-emerald-950/30 border border-emerald-800/30 text-emerald-300 p-3 rounded-lg text-xs flex items-center gap-2">
            <Check className="w-4 h-4" />
            <span>Auto-Purge Triggered: Removed {purgeAlertCount} disposed baggage records older than 7 days from persistent storage.</span>
          </div>
        )}

        {/* 4. Live Dashboard Bento Grid */}
        {activeSection === 'dashboard' && (
          <section id="dashboard" className="space-y-4">
             <div className="flex justify-between items-center">
              <h2 className="text-sm font-bold uppercase tracking-wider text-slate-400">Live Operational Dashboard</h2>
              <div className="flex bg-slate-900 border border-slate-800 rounded-lg p-1">
                <button
                  onClick={() => setActiveRegistry('Arrival')}
                  className={`px-3 py-1 text-xs font-medium rounded-md transition ${activeRegistry === 'Arrival' ? 'bg-indigo-600 text-white' : 'text-slate-400 hover:text-slate-200'}`}
                >
                  Arrival
                </button>
                <button
                  onClick={() => setActiveRegistry('Departure')}
                  className={`px-3 py-1 text-xs font-medium rounded-md transition ${activeRegistry === 'Departure' ? 'bg-indigo-600 text-white' : 'text-slate-400 hover:text-slate-200'}`}
                >
                  Departure
                </button>
                <button
                  onClick={() => setActiveRegistry('Combined')}
                  className={`px-3 py-1 text-xs font-medium rounded-md transition ${activeRegistry === 'Combined' ? 'bg-indigo-600 text-white' : 'text-slate-400 hover:text-slate-200'}`}
                >
                  Combined
                </button>
              </div>
              {activeFilter.type !== 'all' && (
                <button
                  onClick={handleClearFilters}
                  className="flex items-center gap-1 text-xs text-indigo-400 hover:text-indigo-300 font-semibold cursor-pointer"
                >
                  <RotateCcw className="w-3 h-3" />
                  <span>Reset Dashboard Filter</span>
                </button>
              )}
            </div>

            {/* Primary counters */}
            <div className="grid grid-cols-2 lg:grid-cols-5 gap-4">
              <button
                onClick={() => {
                  setActiveFilter({ type: 'all' });
                  setActiveSection('registry');
                }}
                className={`text-left p-4 rounded-xl border transition-all cursor-pointer ${
                  activeFilter.type === 'all'
                    ? 'bg-blue-600/20 border-blue-500 ring-1 ring-blue-500'
                    : 'bg-slate-900 border-slate-800 hover:border-slate-700'
                }`}
              >
                <div className="text-xs font-medium text-slate-400">Expected Manifest</div>
                <div className="text-3xl font-bold text-slate-100 mt-1 font-mono">{stats.totalExpected}</div>
                <div className="text-[10px] text-slate-500 mt-1">Total imported lists</div>
              </button>

              <button
                onClick={() => {
                  setActiveFilter({ type: 'arrived' });
                  setActiveSection('registry');
                }}
                className={`text-left p-4 rounded-xl border transition-all cursor-pointer ${
                  activeFilter.type === 'arrived'
                    ? 'bg-emerald-600/20 border-emerald-500 ring-1 ring-emerald-500'
                    : 'bg-slate-900 border-slate-800 hover:border-slate-700'
                }`}
              >
                <div className="text-xs font-medium text-slate-400">Arrived & Received</div>
                <div className="text-3xl font-bold text-emerald-400 mt-1 font-mono">{stats.arrived}</div>
                <div className="text-[10px] text-slate-500 mt-1">Received at belt/gate</div>
              </button>

              <button
                onClick={() => {
                  setActiveFilter({ type: 'non-arrival' });
                  setActiveSection('registry');
                }}
                className={`text-left p-4 rounded-xl border transition-all cursor-pointer ${
                  activeFilter.type === 'non-arrival'
                    ? 'bg-amber-600/20 border-amber-500 ring-1 ring-amber-500'
                    : 'bg-slate-900 border-slate-800 hover:border-slate-700'
                }`}
              >
                <div className="text-xs font-medium text-slate-400">Non-Arrivals</div>
                <div className="text-3xl font-bold text-amber-400 mt-1 font-mono">{stats.nonArrivals}</div>
                <div className="text-[10px] text-slate-500 mt-1">Expected but pending</div>
              </button>

              <button
                onClick={() => {
                  setActiveFilter({ type: 'cleared' });
                  setActiveSection('registry');
                }}
                className={`text-left p-4 rounded-xl border transition-all cursor-pointer ${
                  activeFilter.type === 'cleared'
                    ? 'bg-indigo-600/20 border-indigo-500 ring-1 ring-indigo-500'
                    : 'bg-slate-900 border-slate-800 hover:border-slate-700'
                }`}
              >
                <div className="text-xs font-medium text-slate-400">Customs Cleared</div>
                <div className="text-3xl font-bold text-indigo-400 mt-1 font-mono">{stats.cleared}</div>
                <div className="text-[10px] text-slate-500 mt-1">Ready for disposition</div>
              </button>

              <button
                onClick={() => {
                  setActiveFilter({ type: 'not-cleared' });
                  setActiveSection('registry');
                }}
                className={`text-left p-4 rounded-xl border transition-all cursor-pointer ${
                  activeFilter.type === 'not-cleared'
                    ? 'bg-red-600/20 border-red-500 ring-1 ring-red-500'
                    : 'bg-slate-900 border-slate-800 hover:border-slate-700'
                }`}
              >
                <div className="text-xs font-medium text-slate-400">Not Cleared / Held</div>
                <div className="text-3xl font-bold text-red-400 mt-1 font-mono">{stats.notCleared}</div>
                <div className="text-[10px] text-slate-500 mt-1">Under custom custody</div>
              </button>
            </div>

            {/* Storage Locations counters */}
            <div className="bg-slate-900/60 border border-slate-800/80 rounded-xl p-4">
              <h3 className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-3">Live Storage & Warehouse Audits</h3>
              <button 
                onClick={findAndResolveDuplicates}
                className="text-[10px] bg-slate-800 hover:bg-slate-700 text-indigo-300 px-2 py-1 rounded mb-3 cursor-pointer"
              >
                Check Duplicates
              </button>
              <div className="grid grid-cols-2 sm:grid-cols-5 gap-3">
                {(Object.keys(stats.locations) as Array<keyof typeof stats.locations>).map(locName => {
                  const isSelected = activeFilter.type === 'location' && activeFilter.value === locName;
                  return (
                    <button
                      key={locName}
                      onClick={() => {
                        setActiveFilter({ type: 'location', value: locName });
                        setActiveSection('registry');
                      }}
                      className={`p-3 rounded-lg border text-left transition cursor-pointer ${
                        isSelected 
                          ? 'bg-indigo-600/20 border-indigo-500 ring-1 ring-indigo-500' 
                          : 'bg-slate-900 border-slate-800/80 hover:border-slate-700'
                      }`}
                    >
                      <div className="text-[10px] text-slate-400 uppercase font-mono tracking-wider">{locName}</div>
                      <div className="text-xl font-bold text-slate-200 mt-0.5 font-mono">
                        {stats.locations[locName]}
                      </div>
                    </button>
                  );
                })}
              </div>
            </div>
          </section>
        )}

        {/* 5. Core Airport Gate Operational Workflow (Standalone Sections) */}
        {activeSection === 'scanner' && (
          <section id="scanner" className="max-w-4xl mx-auto space-y-6 w-full">
            {/* Quick Barcode Scanner Module */}
            <div className="bg-slate-900 border border-slate-800 rounded-xl p-5 shadow-lg relative overflow-hidden">
              <div className="absolute top-0 right-0 w-24 h-24 bg-blue-500/5 rounded-full pointer-events-none translate-x-8 -translate-y-8" />
              
              <div className="flex items-center gap-2 mb-3">
                <QrCode className="w-5 h-5 text-indigo-400" />
                <h3 className="font-bold text-slate-200 text-sm">Gate Scanner & Reconciliation</h3>
              </div>
              <div className="flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-3 mb-4">
                <p className="text-xs text-slate-400">
                  Scan or search via Original Tag, Rush Tag, PIR, or Passenger Name.
                </p>
                <button
                  type="button"
                  onClick={() => {
                    setScannerTargetField('gate');
                    setIsScannerOpen(true);
                  }}
                  className="flex items-center justify-center gap-2 bg-indigo-600 hover:bg-indigo-500 text-white text-[10px] uppercase font-bold tracking-wider px-4 py-2 rounded-lg transition"
                >
                  <Camera className="w-4 h-4" />
                  Scan Baggage Tag
                </button>
              </div>

              <form onSubmit={handleScanSubmit} className="space-y-3">
                <div className="relative">
                  <input
                    type="text"
                    value={scannerInput}
                    onChange={(e) => setScannerInput(e.target.value)}
                    placeholder="Enter Bag Tag (e.g., 0724102943) or PIR"
                    className="w-full bg-slate-950 border border-slate-800 focus:border-indigo-500 text-slate-200 px-4 py-3 rounded-lg text-xs font-mono tracking-wider"
                  />
                  <button
                    type="submit"
                    className="absolute right-2 top-1/2 -translate-y-1/2 bg-indigo-600 hover:bg-indigo-500 text-white text-[10px] uppercase font-bold tracking-wider px-3 py-1.5 rounded cursor-pointer"
                  >
                    Verify
                  </button>
                </div>
              </form>

              {scannerNotification && (
                <div className={`mt-3 p-3 rounded-lg text-xs border ${
                  scannerNotification.type === 'success' 
                    ? 'bg-emerald-950/40 border-emerald-500/30 text-emerald-300'
                    : scannerNotification.type === 'warning'
                    ? 'bg-amber-950/40 border-amber-500/30 text-amber-300'
                    : 'bg-red-950/40 border-red-500/30 text-red-300'
                }`}>
                  <p className="font-semibold">{scannerNotification.text}</p>
                  {scannerNotification.type === 'warning' && scannerSearchType === 'tag' && (
                    <div className="mt-3 flex flex-col sm:flex-row gap-2">
                      <button
                        type="button"
                        onClick={() => {
                          const canonical = getCanonicalTag(scannedQuery, iataAirlineMap);
                          setNewBag(prev => ({
                            ...prev,
                            originalTag: canonical,
                            rushTag: '',
                            receivedAt: new Date().toISOString()
                          }));
                          setRegistrationContext('arrival');
                          setShowAddForm(true);
                          setScannerNotification(null);
                        }}
                        className="bg-indigo-600 hover:bg-indigo-500 text-white text-[10px] font-black uppercase tracking-widest px-4 py-2.5 rounded-xl cursor-pointer transition shadow-xl shadow-indigo-600/10 active:scale-95 text-center"
                      >
                        Register as Original Tag
                      </button>
                      <button
                        type="button"
                        onClick={() => {
                          const canonical = getCanonicalTag(scannedQuery, iataAirlineMap);
                          setNewBag(prev => ({
                            ...prev,
                            originalTag: '',
                            rushTag: canonical,
                            receivedAt: new Date().toISOString()
                          }));
                          setRegistrationContext('arrival');
                          setShowAddForm(true);
                          setScannerNotification(null);
                        }}
                        className="bg-amber-600 hover:bg-amber-500 text-white text-[10px] font-black uppercase tracking-widest px-4 py-2.5 rounded-xl cursor-pointer transition shadow-xl shadow-amber-600/10 active:scale-95 text-center"
                      >
                        Register as Rush Tag
                      </button>
                      <button
                        type="button"
                        onClick={() => setScannerNotification(null)}
                        className="bg-slate-800 hover:bg-slate-700 text-slate-400 text-[10px] font-black uppercase tracking-widest px-4 py-2.5 rounded-xl cursor-pointer transition text-center"
                      >
                        Cancel
                      </button>
                    </div>
                  )}
                </div>
              )}

              {/* SEARCH RESULTS DISPLAY */}
              {(scannerResult || scannerResults.length > 0) && (
                <div className="mt-4 space-y-3 animate-in slide-in-from-top-2 duration-300">
                  <div className="flex items-center justify-between border-b border-slate-800 pb-2">
                    <h4 className="text-[10px] font-bold text-slate-500 uppercase tracking-widest">Search Results ({scannerResult ? 1 : scannerResults.length})</h4>
                    <button type="button" onClick={() => { setScannerResult(null); setScannerResults([]); }} className="text-slate-500 hover:text-slate-300">
                      <X className="w-3 h-3" />
                    </button>
                  </div>
                  
                  <div className="grid gap-3">
                    {(scannerResult ? [scannerResult] : scannerResults).map(record => (
                      <div key={record.id} className="bg-slate-950/60 border border-slate-800 rounded-xl p-4 flex flex-col gap-3">
                        <div className="flex justify-between items-start">
                          <div>
                            <div className="text-xs font-bold text-slate-100">{record.name}</div>
                            <div className="text-[10px] text-slate-500 font-mono">{record.pir || 'No PIR'}</div>
                          </div>
                          <div className={`px-2 py-0.5 rounded text-[9px] font-bold uppercase tracking-tighter ${
                            record.status === 'Received' ? 'bg-emerald-500/10 text-emerald-400 border border-emerald-500/20' : 'bg-amber-500/10 text-amber-400 border border-amber-500/20'
                          }`}>
                            {record.status}
                          </div>
                        </div>
                        
                        <div className="grid grid-cols-2 gap-y-2 gap-x-4 text-[10px]">
                          <div className="flex flex-col">
                            <span className="text-slate-500 uppercase text-[8px] font-bold tracking-tight">Original Tag</span>
                            <span className="text-slate-200 font-mono">{record.originalTag || '-'}</span>
                          </div>
                          <div className="flex flex-col">
                            <span className="text-slate-500 uppercase text-[8px] font-bold tracking-tight">Rush Tag</span>
                            <span className="text-indigo-400 font-mono">{record.rushTag || '-'}</span>
                          </div>
                          <div className="flex flex-col">
                            <span className="text-slate-500 uppercase text-[8px] font-bold tracking-tight">Flight / Dest</span>
                            <span className="text-slate-200">{record.flightNo} / {record.destination}</span>
                          </div>
                          <div className="flex flex-col">
                            <span className="text-slate-500 uppercase text-[8px] font-bold tracking-tight">Storage Location</span>
                            <span className="text-slate-200">{record.dispositionLocation || 'Not Assigned'}</span>
                          </div>
                        </div>

                        <div className="flex gap-2 mt-1 pt-3 border-t border-slate-800/50">
                          {record.status !== 'Received' && (
                            <button
                              type="button"
                              onClick={() => {
                                const updated = baggageList.map(b => b.id === record.id ? {
                                  ...b,
                                  status: 'Received' as const,
                                  receivedAt: new Date().toISOString(),
                                  disposition: 'Storage' as const,
                                  dispositionLocation: 'LHG Office' as const,
                                  dispositionUpdatedAt: new Date().toISOString()
                                } : b);
                                saveBaggageData(updated);
                                setScannerResult(updated.find(b => b.id === record.id) || null);
                                setScannerNotification({ text: 'SUCCESS: Bag marked as RECEIVED.', type: 'success' });
                              }}
                              className="flex-1 bg-emerald-600 hover:bg-emerald-500 text-white px-3 py-1.5 rounded-lg text-[10px] font-bold transition"
                            >
                              Mark Received
                            </button>
                          )}
                          <button
                            type="button"
                            onClick={() => handleOpenEditDialog(record)}
                            className="flex-1 bg-slate-800 hover:bg-slate-700 text-slate-200 px-3 py-1.5 rounded-lg text-[10px] font-bold transition border border-slate-700"
                          >
                            Open Details
                          </button>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>

            {/* Quick Actions Buttons */}
            <div className="flex gap-4">
              <button
                onClick={() => {
                   setRegistrationContext('arrival');
                   setShowAddForm(true);
                }}
                className="flex-1 bg-indigo-600 hover:bg-indigo-500 text-white px-4 py-3 rounded-xl flex items-center justify-center gap-2 text-xs font-bold transition shadow shadow-indigo-600/10 cursor-pointer"
              >
                <PlusCircle className="w-4 h-4" />
                <span>Register Arrival</span>
              </button>

              <button
                onClick={() => {
                   setRegistrationContext('departure');
                   setShowAddForm(true);
                }}
                className="flex-1 bg-indigo-600 hover:bg-indigo-500 text-white px-4 py-3 rounded-xl flex items-center justify-center gap-2 text-xs font-bold transition shadow shadow-indigo-600/10 cursor-pointer"
              >
                <PlusCircle className="w-4 h-4" />
                <span>Register Departure</span>
              </button>
            </div>

            {showAddForm && (
              <div className="bg-slate-900 border border-indigo-500/30 rounded-xl p-5 shadow-xl space-y-4 animate-in fade-in zoom-in duration-200">
                <div className="flex justify-between items-center border-b border-slate-800 pb-2">
                  <div className="flex flex-col">
                    <h3 className="font-bold text-slate-200 text-sm">
                      {registrationContext === 'arrival' ? 'Register Arrival Baggage' : 'Register Departure Baggage'}
                    </h3>
                    <span className="text-[10px] text-slate-400">
                      {registrationContext === 'arrival' ? 'Arrival Left-Behind Registry' : 'Departure Left-Behind Registry'}
                    </span>
                  </div>
                  <div className="flex items-center gap-2">
                    <button onClick={() => setShowAddForm(false)} className="cursor-pointer">
                      <X className="w-4 h-4 text-slate-400 hover:text-slate-200" />
                    </button>
                  </div>
                </div>

                {/* Mode Selector Tab */}
                <div className="flex bg-slate-950 p-1 rounded-lg border border-slate-800">
                  <button
                    type="button"
                    onClick={() => setIsBulkMode(true)}
                    className={`flex-1 py-1.5 text-xs rounded-md font-semibold transition cursor-pointer ${isBulkMode ? 'bg-indigo-600 text-white shadow' : 'text-slate-400 hover:text-slate-200'}`}
                  >
                    Bulk Baggage Operations
                  </button>
                  <button
                    type="button"
                    onClick={() => setIsBulkMode(false)}
                    className={`flex-1 py-1.5 text-xs rounded-md font-semibold transition cursor-pointer ${!isBulkMode ? 'bg-indigo-600 text-white shadow' : 'text-slate-400 hover:text-slate-200'}`}
                  >
                    Single Bag Registration
                  </button>
                </div>

                {!isBulkMode ? (
                  /* SINGLE BAG REGISTRATION MODE */
                  <form onSubmit={handleAddBagSubmit} className="space-y-4">
                    {/* Primary operational information (First 5 fields) */}
                    <div className="grid grid-cols-1 md:grid-cols-5 gap-3 bg-slate-950/40 p-4 rounded-lg border border-slate-800/60">
                      {/* 1. Flight No */}
                      <div>
                        <label className="text-[10px] font-bold text-slate-400 uppercase">Flight No *</label>
                        <select
                          required
                          value={newBag.flightNo}
                          onChange={(e) => setNewBag({...newBag, flightNo: e.target.value})}
                          className="w-full bg-slate-950 border border-slate-800 focus:border-indigo-500 text-slate-200 px-2 py-1.5 rounded text-xs cursor-pointer font-semibold"
                        >
                          <option value="">Select Flight</option>
                          {registrationContext === 'arrival' ? (
                            <>
                              {ARRIVAL_FLIGHTS.map(f => <option key={f} value={f}>{f}</option>)}
                              <option value="LHG Other">LHG Other</option>
                              <option value="Other Airline (OAL) Received Bags">Other Airline (OAL) Received Bags</option>
                            </>
                          ) : (
                            <>
                              {DEPARTURE_FLIGHTS.map(f => <option key={f} value={f}>{f}</option>)}
                              <option value="LHG Other">LHG Other</option>
                            </>
                          )}
                        </select>
                      </div>

                      {/* 2. Original Tag */}
                      <div>
                        <label className="text-[10px] font-bold text-slate-400 uppercase">Original Tag</label>
                        <div className="flex">
                          <input
                            type="text"
                            placeholder="10-digit number"
                            value={newBag.originalTag}
                            onChange={(e) => setNewBag({...newBag, originalTag: e.target.value})}
                            className="flex-1 min-w-0 bg-slate-950 border border-slate-800 focus:border-indigo-500 text-slate-200 px-3 py-1.5 rounded-l text-xs font-mono"
                          />
                          <button
                            type="button"
                            onClick={() => {
                              setScannerTargetField('originalTag');
                              setIsScannerOpen(true);
                            }}
                            className="px-2.5 bg-slate-800 hover:bg-slate-700 border border-l-0 border-slate-800 text-indigo-400 rounded-r flex items-center justify-center cursor-pointer transition"
                            title="Scan Original Tag"
                          >
                            <Camera className="w-4 h-4" />
                          </button>
                        </div>
                      </div>

                      {/* 3. Rush Tag */}
                      <div>
                        <label className="text-[10px] font-bold text-slate-400 uppercase">Rush Tag</label>
                        <div className="flex">
                          <input
                            type="text"
                            placeholder="e.g. LX920394"
                            value={newBag.rushTag}
                            onChange={(e) => setNewBag({...newBag, rushTag: e.target.value})}
                            className="flex-1 min-w-0 bg-slate-950 border border-slate-800 focus:border-indigo-500 text-slate-200 px-3 py-1.5 rounded-l text-xs font-mono"
                          />
                          <button
                            type="button"
                            onClick={() => {
                              setScannerTargetField('rushTag');
                              setIsScannerOpen(true);
                            }}
                            className="px-2.5 bg-slate-800 hover:bg-slate-700 border border-l-0 border-slate-800 text-indigo-400 rounded-r flex items-center justify-center cursor-pointer transition"
                            title="Scan Rush Tag"
                          >
                            <Camera className="w-4 h-4" />
                          </button>
                        </div>
                      </div>

                      {/* 4. Name */}
                      <div className="md:col-span-1">
                        <label className="text-[10px] font-bold text-slate-400 uppercase">Passenger Name</label>
                        <input
                          type="text"
                          placeholder="LASTNAME FIRSTNAME"
                          value={newBag.name}
                          onChange={(e) => setNewBag({...newBag, name: e.target.value})}
                          className="w-full bg-slate-950 border border-slate-800 focus:border-indigo-500 text-slate-200 px-3 py-1.5 rounded text-xs"
                        />
                      </div>

                      {/* Date of Receipt */}
                      <div className="md:col-span-1">
                        <label className="text-[10px] font-bold text-slate-400 uppercase">Receipt Date *</label>
                        <input
                          type="datetime-local"
                          value={newBag.receivedAt ? new Date(new Date(newBag.receivedAt).getTime() - (new Date().getTimezoneOffset() * 60000)).toISOString().slice(0, 16) : ''}
                          onChange={(e) => setNewBag({...newBag, receivedAt: e.target.value ? new Date(e.target.value).toISOString() : undefined})}
                          className="w-full bg-slate-950 border border-slate-800 focus:border-indigo-500 text-slate-200 px-3 py-1.5 rounded text-xs font-mono"
                        />
                      </div>

                      {/* 5. PIR */}
                      <div>
                        <label className="text-[10px] font-bold text-slate-400 uppercase">PIR Number</label>
                        <input
                          type="text"
                          placeholder="e.g. DEL_LX_11029"
                          value={newBag.pir}
                          onChange={(e) => setNewBag({...newBag, pir: e.target.value})}
                          className="w-full bg-slate-950 border border-slate-800 focus:border-indigo-500 text-slate-200 px-3 py-1.5 rounded text-xs font-mono"
                        />
                      </div>
                    </div>

                    {/* Rest of the fields */}
                    <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
                      {/* Locked (LN) free text input */}
                      <div>
                        <label className="text-[10px] font-bold text-slate-400 uppercase block mb-1">Locked (L/N)</label>
                        <input
                          type="text"
                          placeholder="e.g. PAD, CL, Y, N"
                          value={newBag.ln || ''}
                          onChange={(e) => setNewBag({...newBag, ln: e.target.value})}
                          className="w-full bg-slate-950 border border-slate-800 focus:border-indigo-500 text-slate-200 px-3 py-1.5 rounded text-xs font-mono"
                        />
                      </div>

                      {/* Weight Field (kg) */}
                      <div>
                        <label className="text-[10px] font-bold text-slate-400 uppercase block mb-1">Weight (kg)</label>
                        <div className="relative">
                          <input
                            type="number"
                            step="0.1"
                            min="0"
                            placeholder="e.g. 23.4 (Optional)"
                            value={newBag.weight === undefined ? '' : newBag.weight}
                            onChange={(e) => setNewBag({...newBag, weight: e.target.value === '' ? undefined : Number(e.target.value)})}
                            className="w-full bg-slate-950 border border-slate-800 focus:border-indigo-500 text-slate-200 pl-3 pr-8 py-1.5 rounded text-xs"
                          />
                          <span className="absolute right-2.5 top-1/2 -translate-y-1/2 text-[10px] text-slate-500 font-bold">kg</span>
                        </div>
                      </div>

                      {/* Damaged Selector (Y/N) */}
                      <div>
                        <label className="text-[10px] font-bold text-slate-400 uppercase block mb-1">Damaged *</label>
                        <div className="flex gap-2">
                          <button
                            type="button"
                            onClick={() => setNewBag({...newBag, damaged: 'Y'})}
                            className={`flex-1 py-1 px-3 text-xs rounded font-semibold border transition cursor-pointer ${
                              newBag.damaged === 'Y'
                                ? 'bg-amber-600 border-amber-500 text-white shadow'
                                : 'bg-slate-950 border-slate-800 text-slate-400 hover:border-slate-700'
                            }`}
                          >
                            Y
                          </button>
                          <button
                            type="button"
                            onClick={() => setNewBag({...newBag, damaged: 'N'})}
                            className={`flex-1 py-1 px-3 text-xs rounded font-semibold border transition cursor-pointer ${
                              newBag.damaged === 'N' || !newBag.damaged
                                ? 'bg-indigo-600 border-indigo-500 text-white shadow'
                                : 'bg-slate-950 border-slate-800 text-slate-400 hover:border-slate-700'
                            }`}
                          >
                            N
                          </button>
                        </div>
                      </div>

                      {/* Bag Status */}
                      <div>
                        <label className="text-[10px] font-bold text-slate-400 uppercase">Bag Status *</label>
                        <select
                          value={newBag.status}
                          onChange={(e) => setNewBag({...newBag, status: e.target.value as any})}
                          className="w-full bg-slate-950 border border-slate-800 focus:border-indigo-500 text-slate-200 px-2 py-1.5 rounded text-xs font-semibold text-indigo-400 cursor-pointer"
                        >
                          <option value="Expected">Expected (Not Arrived)</option>
                          <option value="Received">Arrived (Received)</option>
                        </select>
                      </div>
                    </div>

                    <div className="grid grid-cols-2 gap-3">
                      {/* Destination */}
                      <div>
                        <label className="text-[10px] font-bold text-slate-400 uppercase">Destination</label>
                        <input
                          type="text"
                          value={newBag.destination}
                          onChange={(e) => setNewBag({...newBag, destination: e.target.value})}
                          className="w-full bg-slate-950 border border-slate-800 focus:border-indigo-500 text-slate-200 px-3 py-1.5 rounded text-xs"
                        />
                      </div>

                      {/* Seal */}
                      <div>
                        <label className="text-[10px] font-bold text-slate-400 uppercase">Seal</label>
                        <input
                          type="text"
                          placeholder="S-xxxxx"
                          value={newBag.seal}
                          onChange={(e) => setNewBag({...newBag, seal: e.target.value})}
                          className="w-full bg-slate-950 border border-slate-800 focus:border-indigo-500 text-slate-200 px-2 py-1.5 rounded text-xs font-mono"
                        />
                      </div>
                    </div>

                    {registrationContext === 'arrival' && newBag.status === 'Received' && (
                      <div className="bg-slate-950 p-3 rounded border border-slate-800 space-y-3">
                        <p className="text-[10px] uppercase font-bold text-indigo-400">Arrived Baggage Quick setup</p>
                        <div className="grid grid-cols-2 gap-2">
                          <div>
                            <label className="text-[9px] uppercase text-slate-400">Customs Status</label>
                            <select
                              value={newBag.customsStatus}
                              onChange={(e) => setNewBag({...newBag, customsStatus: e.target.value as any})}
                              className="w-full bg-slate-900 border border-slate-800 text-slate-200 text-xs py-1 rounded cursor-pointer font-semibold"
                            >
                              <option value="Pending">Pending</option>
                              <option value="Cleared">Cleared</option>
                              <option value="Not Cleared">Not Cleared</option>
                              <option value="Marked Preventive">Marked Preventive (Severe Hold)</option>
                            </select>
                          </div>
                          {(newBag.customsStatus === 'Not Cleared' || newBag.customsStatus === 'Marked Preventive') && (
                            <div>
                              <label className="text-[9px] uppercase text-slate-400">Hold Reason</label>
                              <select
                                value={newBag.customsReason}
                                onChange={(e) => setNewBag({...newBag, customsReason: e.target.value as any})}
                                className="w-full bg-slate-900 border border-slate-800 text-slate-200 text-xs py-1 rounded cursor-pointer font-semibold"
                              >
                                <option value="Lack of documents">Lack of documents</option>
                                <option value="Awaiting documents">Awaiting documents</option>
                                <option value="Refused">Refused</option>
                                <option value="Deferred">Deferred</option>
                                <option value="Preventive">Preventive</option>
                              </select>
                            </div>
                          )}
                        </div>
                      </div>
                    )}

                    {registrationContext === 'departure' && (
                      <div className="space-y-4">
                        <div className="grid grid-cols-1 md:grid-cols-2 gap-4 bg-slate-950/40 p-4 rounded-lg border border-slate-800/60">
                          {/* Forwarding Status */}
                          <div className="space-y-1">
                            <label className="text-[10px] font-bold text-slate-400 uppercase">Forwarding Status *</label>
                            <select
                              required
                              value={newBag.departureForwardingStatus}
                              onChange={(e) => setNewBag({...newBag, departureForwardingStatus: e.target.value as any})}
                              className="w-full bg-slate-950 border border-slate-800 focus:border-indigo-500 text-slate-200 px-3 py-2 rounded text-xs font-semibold cursor-pointer"
                            >
                              <option value="Awaiting Forwarding">Awaiting Forwarding</option>
                              <option value="Forwarding Sent">Forwarding Sent</option>
                            </select>
                          </div>

                          {/* Storage Location */}
                          <div className="space-y-1">
                            <label className="text-[10px] font-bold text-slate-400 uppercase">Storage Location *</label>
                            <select
                              required
                              value={newBag.departureStorageLocation}
                              onChange={(e) => setNewBag({...newBag, departureStorageLocation: e.target.value as any})}
                              className="w-full bg-slate-950 border border-slate-800 focus:border-indigo-500 text-slate-200 px-3 py-2 rounded text-xs font-semibold cursor-pointer"
                            >
                              <option value="LHG Office">LHG Office</option>
                              <option value="BMA">BMA</option>
                              <option value="Level 4">Level 4</option>
                            </select>
                          </div>
                        </div>

                        {/* Conditional Forwarding Sent Fields */}
                        <AnimatePresence>
                          {newBag.departureForwardingStatus === 'Forwarding Sent' && (
                            <motion.div
                              initial={{ opacity: 0, height: 0 }}
                              animate={{ opacity: 1, height: 'auto' }}
                              exit={{ opacity: 0, height: 0 }}
                              className="grid grid-cols-2 md:grid-cols-4 gap-3 bg-indigo-950/10 p-3 rounded-lg border border-indigo-500/20"
                            >
                              <div className="col-span-2 md:col-span-1">
                                <label className="text-[9px] font-bold text-indigo-400 uppercase">Forwarding Flight *</label>
                                <input
                                  type="text"
                                  required
                                  placeholder="e.g. LH763"
                                  value={newBag.forwardingFlightNo || ''}
                                  onChange={(e) => setNewBag({...newBag, forwardingFlightNo: e.target.value.toUpperCase()})}
                                  className="w-full bg-slate-950 border border-slate-800 text-slate-200 px-2 py-1.5 rounded text-xs font-mono"
                                />
                              </div>
                              <div className="col-span-2 md:col-span-1">
                                <label className="text-[9px] font-bold text-indigo-400 uppercase">Forwarding Date *</label>
                                <input
                                  type="date"
                                  required
                                  value={newBag.forwardingDate || ''}
                                  onChange={(e) => setNewBag({...newBag, forwardingDate: e.target.value})}
                                  className="w-full bg-slate-950 border border-slate-800 text-slate-200 px-2 py-1.5 rounded text-xs font-mono"
                                />
                              </div>
                              <div className="col-span-2 md:col-span-1">
                                <label className="text-[9px] font-bold text-indigo-400 uppercase">Destination *</label>
                                <input
                                  type="text"
                                  required
                                  placeholder="e.g. FRA"
                                  value={newBag.forwardingDestination || ''}
                                  onChange={(e) => setNewBag({...newBag, forwardingDestination: e.target.value.toUpperCase()})}
                                  className="w-full bg-slate-950 border border-slate-800 text-slate-200 px-2 py-1.5 rounded text-xs font-mono"
                                />
                              </div>
                              <div className="col-span-2 md:col-span-1">
                                <label className="text-[9px] font-bold text-indigo-400 uppercase">Forwarded By *</label>
                                <input
                                  type="text"
                                  required
                                  placeholder="Staff ID / Name"
                                  value={newBag.forwardedBy || ''}
                                  onChange={(e) => setNewBag({...newBag, forwardedBy: e.target.value.toUpperCase()})}
                                  className="w-full bg-slate-950 border border-slate-800 text-slate-200 px-2 py-1.5 rounded text-xs"
                                />
                              </div>
                              <div className="col-span-2 md:col-span-4">
                                <label className="text-[9px] font-bold text-indigo-400 uppercase">Forwarding Remarks (Optional)</label>
                                <input
                                  type="text"
                                  placeholder="Additional forwarding details..."
                                  value={newBag.forwardingRemarks || ''}
                                  onChange={(e) => setNewBag({...newBag, forwardingRemarks: e.target.value})}
                                  className="w-full bg-slate-950 border border-slate-800 text-slate-200 px-2 py-1.5 rounded text-xs"
                                />
                              </div>
                            </motion.div>
                          )}
                        </AnimatePresence>

                        {/* Conditional Storage Location Fields */}
                        <AnimatePresence>
                          {newBag.departureStorageLocation === 'BMA' && (
                            <motion.div
                              initial={{ opacity: 0, height: 0 }}
                              animate={{ opacity: 1, height: 'auto' }}
                              exit={{ opacity: 0, height: 0 }}
                              className="bg-amber-950/10 p-3 rounded-lg border border-amber-500/20"
                            >
                              <label className="text-[9px] font-bold text-amber-400 uppercase">BMA Storage Location / Remarks *</label>
                              <input
                                type="text"
                                required
                                placeholder="e.g. Cage 3, Rack A2, Shelf B..."
                                value={newBag.bmaStorageLocation || ''}
                                onChange={(e) => setNewBag({...newBag, bmaStorageLocation: e.target.value})}
                                className="w-full bg-slate-950 border border-slate-800 text-slate-200 px-3 py-1.5 rounded text-xs"
                              />
                            </motion.div>
                          )}
                        </AnimatePresence>

                        <AnimatePresence>
                          {(newBag.departureStorageLocation === 'LHG Office' || newBag.departureStorageLocation === 'Level 4') && (
                            <motion.div
                              initial={{ opacity: 0, height: 0 }}
                              animate={{ opacity: 1, height: 'auto' }}
                              exit={{ opacity: 0, height: 0 }}
                              className="bg-slate-950/40 p-3 rounded-lg border border-slate-800"
                            >
                              <label className="text-[9px] font-bold text-slate-400 uppercase">Storage Remarks (Optional)</label>
                              <input
                                type="text"
                                placeholder="Internal storage notes..."
                                value={newBag.storageRemarks || ''}
                                onChange={(e) => setNewBag({...newBag, storageRemarks: e.target.value})}
                                className="w-full bg-slate-950 border border-slate-800 text-slate-200 px-3 py-1.5 rounded text-xs"
                              />
                            </motion.div>
                          )}
                        </AnimatePresence>
                      </div>
                    )}

                    {/* Dispositions & Customs Operations Protocol Workflow Selector - Arrival Only */}
                    {registrationContext === 'arrival' && (
                      <>
                        <div className="space-y-1 bg-slate-950/20 p-4 rounded-lg border border-slate-800/50">
                          <label className="text-[10px] font-bold text-indigo-400 uppercase block mb-1">Dispositions &amp; Customs Operations Protocol *</label>
                          <select
                            value={newBag.protocol}
                            onChange={(e) => {
                              const nextProtocol = e.target.value as any;
                              setNewBagClearedAction('');
                              setNewBagNonClearedAction('');
                              setNewBag({
                                ...newBag,
                                protocol: nextProtocol,
                                deliveryAgent: undefined,
                                storageOption: undefined,
                                domesticForwarding: undefined,
                                arrivalBelt: undefined,
                                handoverOption: undefined,
                                warehouseOption: undefined,
                                reexportOption: undefined
                              });
                            }}
                            className="w-full bg-slate-950 border border-slate-800 focus:border-indigo-500 text-slate-200 px-3 py-2 rounded text-xs font-semibold cursor-pointer"
                          >
                            <option value="">-- Choose Protocol to determine remainder of workflow --</option>
                            <option value="Cleared Baggage">Cleared Baggage</option>
                            <option value="Non-Cleared / Other">Non-Cleared / Other</option>
                          </select>
                        </div>

                        {/* Dynamic Fields with smooth Framer Motion expand/collapse animation */}
                        <AnimatePresence initial={false}>
                          {newBag.protocol && (
                            <motion.div
                              key={newBag.protocol}
                              initial={{ opacity: 0, height: 0, scale: 0.98 }}
                              animate={{ opacity: 1, height: 'auto', scale: 1 }}
                              exit={{ opacity: 0, height: 0, scale: 0.98 }}
                              transition={{ duration: 0.3, ease: 'easeInOut' }}
                              className="bg-slate-950/80 p-4 rounded-lg border border-slate-800 space-y-4 overflow-hidden"
                            >
                              <div className="border-b border-slate-800/80 pb-2 flex items-center justify-between">
                                <span className="text-[11px] font-bold text-indigo-400 uppercase tracking-wider">
                                  {newBag.protocol} Workflow Protocol
                                </span>
                                <span className="text-[9px] bg-slate-800 px-2 py-0.5 rounded text-slate-400">Rule-based options</span>
                              </div>

                              {newBag.protocol === 'Cleared Baggage' ? (
                                <div className="space-y-4">
                                  {/* Level 2: Master dropdown for Cleared Baggage Action */}
                                  <div className="space-y-1">
                                    <label className="text-[10px] font-bold text-slate-300 uppercase block">Cleared Baggage Action *</label>
                                    <select
                                      value={newBagClearedAction}
                                      onChange={(e) => {
                                        const act = e.target.value;
                                        setNewBagClearedAction(act);
                                        setNewBag(prev => ({
                                          ...prev,
                                          deliveryAgent: act === 'deliveryAgent' ? 'VVM' : undefined,
                                          storageOption: act === 'storage' ? 'Standard Warehousing – LHG Office' : undefined,
                                          domesticForwarding: act === 'domesticForwarding' ? 'Air India' : undefined
                                        }));
                                      }}
                                      className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-1.5 px-3 rounded cursor-pointer font-semibold"
                                    >
                                      <option value="">-- Select Cleared Baggage Action --</option>
                                      <option value="deliveryAgent">Delivery Agent</option>
                                      <option value="storage">Storage</option>
                                      <option value="domesticForwarding">Domestic Baggage Forwarding</option>
                                    </select>
                                  </div>

                                  <AnimatePresence initial={false}>
                                    {newBagClearedAction === 'deliveryAgent' && (
                                      <motion.div
                                        initial={{ opacity: 0, height: 0 }}
                                        animate={{ opacity: 1, height: 'auto' }}
                                        exit={{ opacity: 0, height: 0 }}
                                        transition={{ duration: 0.2 }}
                                        className="space-y-1 overflow-hidden"
                                      >
                                        <label className="text-[10px] font-bold text-slate-300 uppercase block">Delivery Agent *</label>
                                        <select
                                          value={newBag.deliveryAgent || 'VVM'}
                                          onChange={(e) => setNewBag({...newBag, deliveryAgent: e.target.value as any})}
                                          className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-1.5 px-3 rounded cursor-pointer"
                                        >
                                          <option value="VVM">VVM</option>
                                          <option value="Outlook">Outlook</option>
                                          <option value="Advik">Advik</option>
                                        </select>
                                      </motion.div>
                                    )}

                                    {newBagClearedAction === 'storage' && (
                                      <motion.div
                                        initial={{ opacity: 0, height: 0 }}
                                        animate={{ opacity: 1, height: 'auto' }}
                                        exit={{ opacity: 0, height: 0 }}
                                        transition={{ duration: 0.2 }}
                                        className="space-y-1 overflow-hidden"
                                      >
                                        <label className="text-[10px] font-bold text-slate-300 uppercase block font-sans">Storage Location *</label>
                                        <select
                                          value={newBag.storageOption || 'Standard Warehousing – LHG Office'}
                                          onChange={(e) => setNewBag({...newBag, storageOption: e.target.value as any})}
                                          className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-1.5 px-3 rounded cursor-pointer"
                                        >
                                          <option value="Standard Warehousing – LHG Office">Standard Warehousing – LHG Office</option>
                                        </select>
                                      </motion.div>
                                    )}

                                    {newBagClearedAction === 'domesticForwarding' && (
                                      <motion.div
                                        initial={{ opacity: 0, height: 0 }}
                                        animate={{ opacity: 1, height: 'auto' }}
                                        exit={{ opacity: 0, height: 0 }}
                                        transition={{ duration: 0.2 }}
                                        className="space-y-1 overflow-hidden"
                                      >
                                        <label className="text-[10px] font-bold text-slate-300 uppercase block">Forward Via *</label>
                                        <select
                                          value={newBag.domesticForwarding || 'Air India'}
                                          onChange={(e) => setNewBag({...newBag, domesticForwarding: e.target.value as any})}
                                          className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-1.5 px-3 rounded cursor-pointer"
                                        >
                                          <option value="Air India">Air India</option>
                                          <option value="IndiGo">IndiGo</option>
                                          <option value="SpiceJet">SpiceJet</option>
                                        </select>
                                      </motion.div>
                                    )}
                                  </AnimatePresence>
                                </div>
                              ) : (
                                <div className="space-y-4">
                                  {/* Level 2: Master dropdown for Non-Cleared / Other Action */}
                                  <div className="space-y-1">
                                    <label className="text-[10px] font-bold text-slate-300 uppercase block">Non-Cleared / Other Action *</label>
                                    <select
                                      value={newBagNonClearedAction}
                                      onChange={(e) => {
                                        const act = e.target.value;
                                        setNewBagNonClearedAction(act);
                                        setNewBag(prev => ({
                                          ...prev,
                                          arrivalBelt: act === 'arrivalBelt' ? 'Arrival Belt 9' : undefined,
                                          handoverOption: act === 'handover' ? 'Partner Airlines' : undefined,
                                          warehouseOption: act === 'warehouse' ? 'CWC Warehouse' : undefined,
                                          reexportOption: act === 'reexport' ? 'Re-export to Carrier Hub' : undefined
                                        }));
                                      }}
                                      className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-1.5 px-3 rounded cursor-pointer font-semibold"
                                    >
                                      <option value="">-- Select Non-Cleared / Other Action --</option>
                                      <option value="arrivalBelt">Arrival Belt</option>
                                      <option value="handover">Handover</option>
                                      <option value="warehouse">Warehouse</option>
                                      <option value="reexport">Re-export</option>
                                    </select>
                                  </div>

                                  <AnimatePresence initial={false}>
                                    {newBagNonClearedAction === 'arrivalBelt' && (
                                      <motion.div
                                        initial={{ opacity: 0, height: 0 }}
                                        animate={{ opacity: 1, height: 'auto' }}
                                        exit={{ opacity: 0, height: 0 }}
                                        transition={{ duration: 0.2 }}
                                        className="space-y-1 overflow-hidden"
                                      >
                                        <label className="text-[10px] font-bold text-slate-300 uppercase block">Arrival Belt *</label>
                                        <select
                                          value={newBag.arrivalBelt || 'Arrival Belt 9'}
                                          onChange={(e) => setNewBag({...newBag, arrivalBelt: e.target.value as any})}
                                          className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-1.5 px-3 rounded cursor-pointer"
                                        >
                                          <option value="Arrival Belt 9">Belt 9 (Default)</option>
                                        </select>
                                        <p className="text-[9px] text-slate-500 italic mt-0.5">Default holding area with queue check.</p>
                                      </motion.div>
                                    )}

                                    {newBagNonClearedAction === 'handover' && (
                                      <motion.div
                                        initial={{ opacity: 0, height: 0 }}
                                        animate={{ opacity: 1, height: 'auto' }}
                                        exit={{ opacity: 0, height: 0 }}
                                        transition={{ duration: 0.2 }}
                                        className="space-y-1 overflow-hidden"
                                      >
                                        <label className="text-[10px] font-bold text-slate-300 uppercase block">Handover To *</label>
                                        <select
                                          value={newBag.handoverOption || 'Partner Airlines'}
                                          onChange={(e) => setNewBag({...newBag, handoverOption: e.target.value as any})}
                                          className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-1.5 px-3 rounded cursor-pointer"
                                        >
                                          <option value="Partner Airlines">Partner Airlines</option>
                                        </select>
                                        <p className="text-[9px] text-slate-500 italic mt-0.5">Transfer custody to the designated partner airline.</p>
                                      </motion.div>
                                    )}

                                    {newBagNonClearedAction === 'warehouse' && (
                                      <motion.div
                                        initial={{ opacity: 0, height: 0 }}
                                        animate={{ opacity: 1, height: 'auto' }}
                                        exit={{ opacity: 0, height: 0 }}
                                        transition={{ duration: 0.2 }}
                                        className="space-y-1 overflow-hidden"
                                      >
                                        <label className="text-[10px] font-bold text-slate-300 uppercase block">Warehouse *</label>
                                        <select
                                          value={newBag.warehouseOption || 'CWC Warehouse'}
                                          onChange={(e) => setNewBag({...newBag, warehouseOption: e.target.value as any})}
                                          className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-1.5 px-3 rounded cursor-pointer"
                                        >
                                          <option value="CWC Warehouse">CWC Warehouse</option>
                                        </select>
                                        <p className="text-[9px] text-slate-500 italic mt-0.5">Secure central depot storage.</p>
                                      </motion.div>
                                    )}

                                    {newBagNonClearedAction === 'reexport' && (
                                      <motion.div
                                        initial={{ opacity: 0, height: 0 }}
                                        animate={{ opacity: 1, height: 'auto' }}
                                        exit={{ opacity: 0, height: 0 }}
                                        transition={{ duration: 0.2 }}
                                        className="space-y-1 overflow-hidden"
                                      >
                                        <label className="text-[10px] font-bold text-slate-300 uppercase block">Re-export Destination *</label>
                                        <select
                                          value={newBag.reexportOption || 'Re-export to Carrier Hub'}
                                          onChange={(e) => setNewBag({...newBag, reexportOption: e.target.value as any})}
                                          className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-1.5 px-3 rounded cursor-pointer"
                                        >
                                          <option value="Re-export to Carrier Hub">Return to Carrier Hub</option>
                                        </select>
                                        <p className="text-[9px] text-slate-500 italic mt-0.5">Repatriate the baggage to the originating carrier&apos;s hub.</p>
                                      </motion.div>
                                    )}
                                  </AnimatePresence>
                                </div>
                              )}
                            </motion.div>
                          )}
                        </AnimatePresence>
                      </>
                    )}

                    {/* Remarks / Left Behind Reason */}
                    <div>
                      <label className="text-[10px] font-bold text-slate-400 uppercase">
                        {registrationContext === 'arrival' ? 'Remarks' : 'Left Behind Reason'}
                      </label>
                      <textarea
                        rows={2}
                        placeholder={registrationContext === 'arrival' ? 'Add handling details...' : 'Late acceptance, Security hold, etc.'}
                        value={registrationContext === 'arrival' ? newBag.remarks : newBag.leftBehindReason}
                        onChange={(e) => {
                          if (registrationContext === 'arrival') {
                            setNewBag({...newBag, remarks: e.target.value});
                          } else {
                            setNewBag({...newBag, leftBehindReason: e.target.value});
                          }
                        }}
                        className="w-full bg-slate-950 border border-slate-800 focus:border-indigo-500 text-slate-200 px-3 py-1.5 rounded text-xs transition-all duration-200 ease-in-out"
                      />
                    </div>

                    <button
                      type="submit"
                      className="w-full bg-gradient-to-r from-indigo-600 to-indigo-500 hover:from-indigo-500 hover:to-indigo-400 text-white font-bold py-2.5 rounded-lg text-xs cursor-pointer shadow-lg shadow-indigo-600/10 transition"
                    >
                      Register and Save Bag
                    </button>
                  </form>
                ) : (
                  /* BULK BAGGAGE OPERATIONS MODE */
                  <div className="space-y-6">
                    {/* 1. Bulk Import Operations */}
                    {registrationContext === 'arrival' && (
                      <div className="space-y-3">
                        <div className="flex items-center gap-2 mb-1">
                          <Upload className="w-4 h-4 text-indigo-400" />
                          <h4 className="text-[11px] font-bold text-slate-300 uppercase tracking-wider">Excel / List Imports</h4>
                        </div>
                        <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
                          <div className="flex flex-col gap-2">
                            <button
                              onClick={() => {
                                setIsSBHImport(false);
                                setIsAlteaImport(false);
                                setShowImportDialog(true);
                              }}
                              className="flex-1 flex flex-col items-center justify-center gap-2 bg-indigo-600/10 border border-indigo-500/20 hover:bg-indigo-600/20 text-indigo-400 text-xs font-bold py-4 rounded-xl transition cursor-pointer group"
                            >
                              <div className="bg-indigo-600/20 p-2 rounded-full group-hover:scale-110 transition-transform">
                                <Upload className="w-5 h-5" />
                              </div>
                              Import BDO Excel
                            </button>
                            {isMobile && (
                              <button
                                onClick={() => {
                                  setOcrTargetMode('table');
                                  cameraInputRef.current?.click();
                                }}
                                className="flex items-center justify-center gap-2 bg-indigo-600/20 hover:bg-indigo-600/30 text-indigo-400 text-[10px] font-bold py-2 rounded-lg transition border border-indigo-500/10"
                              >
                                <Camera className="w-3 h-3" />
                                Capture Sheet
                              </button>
                            )}
                          </div>

                          <div className="flex flex-col gap-2">
                            <button
                              onClick={() => {
                                setIsAlteaImport(false);
                                setShowSBHInstructions(true);
                              }}
                              className="flex-1 flex flex-col items-center justify-center gap-2 bg-amber-600/10 border border-amber-500/20 hover:bg-amber-600/20 text-amber-400 text-xs font-bold py-4 rounded-xl transition cursor-pointer group"
                            >
                              <div className="bg-amber-600/20 p-2 rounded-full group-hover:scale-110 transition-transform">
                                <FileSpreadsheet className="w-5 h-5" />
                              </div>
                              Import Rush Bag List (SBH)
                            </button>
                            {isMobile && (
                              <button
                                onClick={() => {
                                  setOcrTargetMode('table');
                                  cameraInputRef.current?.click();
                                }}
                                className="flex items-center justify-center gap-2 bg-amber-600/20 hover:bg-amber-600/30 text-amber-400 text-[10px] font-bold py-2 rounded-lg transition border border-amber-500/10"
                              >
                                <Camera className="w-3 h-3" />
                                Capture Sheet
                              </button>
                            )}
                          </div>

                          <div className="flex flex-col gap-2">
                            <button
                              onClick={() => {
                                setIsSBHImport(false);
                                setIsAlteaImport(true);
                                setShowAlteaInstructions(true);
                              }}
                              className="flex-1 flex flex-col items-center justify-center gap-2 bg-emerald-600/10 border border-emerald-500/20 hover:bg-emerald-600/20 text-emerald-400 text-xs font-bold py-4 rounded-xl transition cursor-pointer group"
                            >
                              <div className="bg-emerald-600/20 p-2 rounded-full group-hover:scale-110 transition-transform">
                                <ClipboardList className="w-5 h-5" />
                              </div>
                              Import Altea Rush Bag List
                            </button>
                            {isMobile && (
                              <button
                                onClick={() => {
                                  setOcrTargetMode('altea');
                                  cameraInputRef.current?.click();
                                }}
                                className="flex items-center justify-center gap-2 bg-emerald-600/20 hover:bg-emerald-600/30 text-emerald-400 text-[10px] font-bold py-2 rounded-lg transition border border-emerald-500/10"
                              >
                                <Camera className="w-3 h-3" />
                                Capture Screen
                              </button>
                            )}
                          </div>
                        </div>
                      </div>
                    )}

                    {/* 2. Bulk Manual Entry & Scanning */}
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                      {/* Bulk Manual Entry Section */}
                      <div className="space-y-3 bg-slate-950/40 p-4 rounded-xl border border-slate-800/60 shadow-inner">
                        <div className="flex items-center gap-2 mb-1">
                          <Keyboard className="w-4 h-4 text-emerald-400" />
                          <h4 className="text-[11px] font-bold text-slate-300 uppercase tracking-wider">Bulk Manual Entry</h4>
                        </div>
                        <div>
                          <label className="text-[11px] font-bold text-slate-300 uppercase block">
                            Baggage Tags List *
                          </label>
                          <span className="text-[9px] text-slate-500">
                            Enter multiple 10-digit tags separated by commas, spaces, or lines.
                          </span>
                        </div>
                        <textarea
                          rows={4}
                          placeholder="Type or paste multiple 10-digit tags here..."
                          value={bulkTagsInput}
                          onChange={(e) => setBulkTagsInput(e.target.value)}
                          className="w-full bg-slate-950 border border-slate-800 focus:border-indigo-500 text-slate-200 px-3 py-2 rounded text-xs font-mono tracking-widest placeholder-slate-700 mt-2"
                        />
                      </div>

                      {/* Bulk Barcode Scanning Section */}
                      <div className="space-y-3 bg-slate-950/40 p-4 rounded-xl border border-slate-800/60 shadow-inner flex flex-col justify-between">
                        <div className="space-y-3">
                          <div className="flex items-center gap-2 mb-1">
                            <QrCode className="w-4 h-4 text-indigo-400" />
                            <h4 className="text-[11px] font-bold text-slate-300 uppercase tracking-wider">Bulk Barcode Scanning</h4>
                          </div>
                          <p className="text-[10px] text-slate-400">
                            Use your camera to scan multiple baggage tags in a continuous session. All detected tags will be added to the list.
                          </p>
                        </div>
                        
                        <button
                          type="button"
                          onClick={() => {
                            setScannerTargetField('bulk');
                            setContinuousScannedTags([]);
                            setIsScannerOpen(true);
                          }}
                          className="w-full py-4 bg-indigo-600 hover:bg-indigo-500 text-white rounded-lg text-xs font-bold transition flex items-center justify-center gap-2 cursor-pointer shadow-lg shadow-indigo-600/20 mt-4"
                        >
                          <Camera className="w-5 h-5" />
                          <span>Launch Bulk Barcode Scanner</span>
                        </button>
                      </div>
                    </div>

                    {/* Parser Stats Preview */}
                    {bulkTagsInput.trim() && (
                      <div className="flex justify-between items-center bg-slate-950/60 px-4 py-2 rounded-lg border border-slate-800 text-[11px] text-slate-400 font-mono">
                        <span>Total Tags Detected for Processing:</span>
                        <span className="font-bold text-emerald-400 text-sm">{parseBulkInput(bulkTagsInput).length} Bags</span>
                      </div>
                    )}

                    <form onSubmit={handleProcessBulkAdd} className="space-y-4">
                    {/* Bulk Dispositions & Customs Operations Protocol */}
                    {registrationContext === 'arrival' ? (
                      <div className="bg-indigo-950/25 p-4 rounded-lg border border-indigo-500/20 space-y-3">
                        <div>
                          <label className="text-[10px] font-bold text-indigo-400 uppercase block mb-1">
                            Dispositions &amp; Customs Operations Protocol *
                          </label>
                          <select
                            required
                            value={bulkProtocol}
                            onChange={(e) => {
                              const nextProtocol = e.target.value as any;
                              setBulkClearedAction('');
                              setBulkNonClearedAction('');
                              setBulkProtocol(nextProtocol);
                            }}
                            className="w-full bg-slate-950 border border-slate-800 focus:border-indigo-500 text-indigo-200 px-3 py-2 rounded text-xs font-semibold cursor-pointer"
                          >
                            <option value="">-- Choose Protocol to determine remainder of workflow --</option>
                            <option value="Cleared Baggage">Cleared Baggage (Cleared for Delivery/Storage/Forwarding)</option>
                            <option value="Non-Cleared / Other">Non-Cleared / Other (Arrival Belts/Handovers/CWC/Re-export)</option>
                          </select>
                        </div>

                        {bulkProtocol && (
                          <div className="bg-slate-950/80 p-3 rounded border border-slate-800 space-y-3">
                            {bulkProtocol === 'Cleared Baggage' ? (
                              <div className="space-y-3">
                                <div className="space-y-1">
                                  <label className="text-[10px] font-bold text-slate-300 uppercase block">Cleared Baggage Action *</label>
                                  <select
                                    required
                                    value={bulkClearedAction}
                                    onChange={(e) => setBulkClearedAction(e.target.value)}
                                    className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-1.5 px-3 rounded cursor-pointer font-semibold"
                                  >
                                    <option value="">-- Select Cleared Baggage Action --</option>
                                    <option value="deliveryAgent">Delivery Agent</option>
                                    <option value="storage">Storage</option>
                                    <option value="domesticForwarding">Domestic Baggage Forwarding</option>
                                  </select>
                                </div>

                                {bulkClearedAction === 'deliveryAgent' && (
                                  <div className="space-y-1">
                                    <label className="text-[10px] font-bold text-slate-300 uppercase block">Delivery Agent *</label>
                                    <select
                                      required
                                      value={bulkDeliveryAgent}
                                      onChange={(e) => setBulkDeliveryAgent(e.target.value as any)}
                                      className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-1.5 px-3 rounded cursor-pointer"
                                    >
                                      <option value="VVM">VVM</option>
                                      <option value="Outlook">Outlook</option>
                                      <option value="Advik">Advik</option>
                                    </select>
                                  </div>
                                )}

                                {bulkClearedAction === 'storage' && (
                                  <div className="space-y-1">
                                    <label className="text-[10px] font-bold text-slate-300 uppercase block">Storage Location *</label>
                                    <select
                                      required
                                      value={bulkStorageOption}
                                      onChange={(e) => setBulkStorageOption(e.target.value as any)}
                                      className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-1.5 px-3 rounded cursor-pointer"
                                    >
                                      <option value="Standard Warehousing – LHG Office">Standard Warehousing – LHG Office</option>
                                    </select>
                                  </div>
                                )}

                                {bulkClearedAction === 'domesticForwarding' && (
                                  <div className="space-y-3">
                                    <div className="space-y-1">
                                      <label className="text-[10px] font-bold text-slate-300 uppercase block">Forward Via *</label>
                                      <select
                                        required
                                        value={bulkDomesticForwarding}
                                        onChange={(e) => setBulkDomesticForwarding(e.target.value as any)}
                                        className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-1.5 px-3 rounded cursor-pointer"
                                      >
                                        <option value="Air India">Air India</option>
                                        <option value="IndiGo">IndiGo</option>
                                        <option value="SpiceJet">SpiceJet</option>
                                      </select>
                                    </div>
                                    <div className="space-y-1">
                                      <label className="text-[10px] font-bold text-slate-300 uppercase block">Forwarding Flight No *</label>
                                      <input
                                        type="text"
                                        required
                                        value={bulkForwardingFlight}
                                        onChange={(e) => setBulkForwardingFlight(e.target.value)}
                                        className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-1.5 px-3 rounded"
                                      />
                                    </div>
                                    <div className="space-y-1">
                                      <label className="text-[10px] font-bold text-slate-300 uppercase block">Forwarding Date *</label>
                                      <input
                                        type="date"
                                        required
                                        value={bulkForwardingDate}
                                        onChange={(e) => setBulkForwardingDate(e.target.value)}
                                        className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-1.5 px-3 rounded"
                                      />
                                    </div>
                                  </div>
                                )}
                              </div>
                            ) : (
                              <div className="space-y-3">
                                <div className="space-y-1">
                                  <label className="text-[10px] font-bold text-slate-300 uppercase block">Non-Cleared / Other Action *</label>
                                  <select
                                    required
                                    value={bulkNonClearedAction}
                                    onChange={(e) => setBulkNonClearedAction(e.target.value)}
                                    className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-1.5 px-3 rounded cursor-pointer font-semibold"
                                  >
                                    <option value="">-- Select Non-Cleared / Other Action --</option>
                                    <option value="arrivalBelt">Arrival Belt</option>
                                    <option value="handover">Handover</option>
                                    <option value="warehouse">Warehouse</option>
                                    <option value="reexport">Re-export</option>
                                  </select>
                                </div>

                                {bulkNonClearedAction === 'arrivalBelt' && (
                                  <div className="space-y-1">
                                    <label className="text-[10px] font-bold text-slate-300 uppercase block">Arrival Belt *</label>
                                    <select
                                      required
                                      value={bulkArrivalBelt}
                                      onChange={(e) => setBulkArrivalBelt(e.target.value as any)}
                                      className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-1.5 px-3 rounded cursor-pointer"
                                    >
                                      <option value="Arrival Belt 9">Belt 9 (Default)</option>
                                    </select>
                                  </div>
                                )}

                                {bulkNonClearedAction === 'handover' && (
                                  <div className="space-y-1">
                                    <label className="text-[10px] font-bold text-slate-300 uppercase block">Handover To *</label>
                                    <select
                                      required
                                      value={bulkHandoverOption}
                                      onChange={(e) => setBulkHandoverOption(e.target.value as any)}
                                      className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-1.5 px-3 rounded cursor-pointer"
                                    >
                                      <option value="Partner Airlines">Partner Airlines</option>
                                    </select>
                                  </div>
                                )}

                                {bulkNonClearedAction === 'warehouse' && (
                                  <div className="space-y-1">
                                    <label className="text-[10px] font-bold text-slate-300 uppercase block">Warehouse *</label>
                                    <select
                                      required
                                      value={bulkWarehouseOption}
                                      onChange={(e) => setBulkWarehouseOption(e.target.value as any)}
                                      className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-1.5 px-3 rounded cursor-pointer"
                                    >
                                      <option value="CWC Warehouse">CWC Warehouse</option>
                                    </select>
                                  </div>
                                )}

                                {bulkNonClearedAction === 'reexport' && (
                                  <div className="space-y-1">
                                    <label className="text-[10px] font-bold text-slate-300 uppercase block">Re-export Destination *</label>
                                    <select
                                      required
                                      value={bulkReexportOption}
                                      onChange={(e) => setBulkReexportOption(e.target.value as any)}
                                      className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-1.5 px-3 rounded cursor-pointer"
                                    >
                                      <option value="Re-export to Carrier Hub">Carrier Hub</option>
                                    </select>
                                  </div>
                                )}
                              </div>
                            )}
                          </div>
                        )}
                      </div>
                    ) : (
                      <div className="bg-indigo-950/25 p-4 rounded-lg border border-indigo-500/20 space-y-4">
                        <div className="grid grid-cols-2 gap-4">
                          {/* Bulk Forwarding Status */}
                          <div className="space-y-1">
                            <label className="text-[10px] font-bold text-indigo-400 uppercase block">Forwarding Status *</label>
                            <select
                              required
                              value={bulkDepartureForwardingStatus}
                              onChange={(e) => setBulkDepartureForwardingStatus(e.target.value as any)}
                              className="w-full bg-slate-950 border border-slate-800 focus:border-indigo-500 text-slate-200 px-3 py-2 rounded text-xs font-semibold cursor-pointer"
                            >
                              <option value="Awaiting Forwarding">Awaiting Forwarding</option>
                              <option value="Forwarding Sent">Forwarding Sent</option>
                            </select>
                          </div>

                          {/* Bulk Storage Location */}
                          <div className="space-y-1">
                            <label className="text-[10px] font-bold text-indigo-400 uppercase block">Storage Location *</label>
                            <select
                              required
                              value={bulkDepartureStorageLocation}
                              onChange={(e) => setBulkDepartureStorageLocation(e.target.value as any)}
                              className="w-full bg-slate-950 border border-slate-800 focus:border-indigo-500 text-slate-200 px-3 py-2 rounded text-xs font-semibold cursor-pointer"
                            >
                              <option value="LHG Office">LHG Office</option>
                              <option value="BMA">BMA</option>
                              <option value="Level 4">Level 4</option>
                            </select>
                          </div>
                        </div>

                        {/* Bulk Conditional Forwarding Fields */}
                        <AnimatePresence>
                          {bulkDepartureForwardingStatus === 'Forwarding Sent' && (
                            <motion.div
                              initial={{ opacity: 0, height: 0 }}
                              animate={{ opacity: 1, height: 'auto' }}
                              exit={{ opacity: 0, height: 0 }}
                              className="grid grid-cols-2 gap-3 bg-slate-950/60 p-3 rounded border border-slate-800"
                            >
                              <div>
                                <label className="text-[9px] font-bold text-slate-400 uppercase">Forwarding Flight *</label>
                                <input
                                  type="text"
                                  required
                                  placeholder="LH763"
                                  value={bulkForwardingFlight}
                                  onChange={(e) => setBulkForwardingFlight(e.target.value.toUpperCase())}
                                  className="w-full bg-slate-900 border border-slate-800 text-slate-200 px-2 py-1.5 rounded text-xs"
                                />
                              </div>
                              <div>
                                <label className="text-[9px] font-bold text-slate-400 uppercase">Forwarding Date *</label>
                                <input
                                  type="date"
                                  required
                                  value={bulkForwardingDate}
                                  onChange={(e) => setBulkForwardingDate(e.target.value)}
                                  className="w-full bg-slate-900 border border-slate-800 text-slate-200 px-2 py-1.5 rounded text-xs"
                                />
                              </div>
                              <div>
                                <label className="text-[9px] font-bold text-slate-400 uppercase">Destination *</label>
                                <input
                                  type="text"
                                  required
                                  placeholder="FRA"
                                  value={bulkForwardingDestination}
                                  onChange={(e) => setBulkForwardingDestination(e.target.value.toUpperCase())}
                                  className="w-full bg-slate-900 border border-slate-800 text-slate-200 px-2 py-1.5 rounded text-xs"
                                />
                              </div>
                              <div>
                                <label className="text-[9px] font-bold text-slate-400 uppercase">Forwarded By *</label>
                                <input
                                  type="text"
                                  required
                                  placeholder="Staff ID"
                                  value={bulkForwardedBy}
                                  onChange={(e) => setBulkForwardedBy(e.target.value.toUpperCase())}
                                  className="w-full bg-slate-900 border border-slate-800 text-slate-200 px-2 py-1.5 rounded text-xs"
                                />
                              </div>
                            </motion.div>
                          )}
                        </AnimatePresence>

                        {/* Bulk Conditional Storage Fields */}
                        <AnimatePresence>
                          {bulkDepartureStorageLocation === 'BMA' && (
                            <motion.div
                              initial={{ opacity: 0, height: 0 }}
                              animate={{ opacity: 1, height: 'auto' }}
                              exit={{ opacity: 0, height: 0 }}
                              className="bg-slate-950/60 p-3 rounded border border-slate-800"
                            >
                              <label className="text-[9px] font-bold text-slate-400 uppercase">BMA Storage Location *</label>
                              <input
                                type="text"
                                required
                                placeholder="Cage/Rack location..."
                                value={bulkBmaStorageLocation}
                                onChange={(e) => setBulkBmaStorageLocation(e.target.value)}
                                className="w-full bg-slate-900 border border-slate-800 text-slate-200 px-2 py-1.5 rounded text-xs"
                              />
                            </motion.div>
                          )}
                        </AnimatePresence>
                      </div>
                    )}
                    
                    {/* Common Baggage Metadata Optional Fields Grid */}
                    <div className="grid grid-cols-2 md:grid-cols-4 gap-3 bg-slate-950/40 p-4 rounded-lg border border-slate-800/60">
                      {/* Flight No */}
                      <div>
                        <label className="text-[10px] font-bold text-slate-400 uppercase">Flight No</label>
                        <select
                          value={bulkFlightNo}
                          onChange={(e) => setBulkFlightNo(e.target.value)}
                          className="w-full bg-slate-950 border border-slate-800 text-slate-200 px-2 py-1.5 rounded text-xs cursor-pointer font-semibold"
                        >
                          <option value="">Select Flight</option>
                          {registrationContext === 'arrival' ? (
                            <>
                              {ARRIVAL_FLIGHTS.map(f => <option key={f} value={f}>{f}</option>)}
                              <option value="LHG Other">LHG Other</option>
                              <option value="Other Airline (OAL) Received Bags">Other Airline (OAL) Received Bags</option>
                            </>
                          ) : (
                            <>
                              {DEPARTURE_FLIGHTS.map(f => <option key={f} value={f}>{f}</option>)}
                              <option value="LHG Other">LHG Other</option>
                            </>
                          )}
                        </select>
                      </div>

                      {/* Destination */}
                      <div>
                        <label className="text-[10px] font-bold text-slate-400 uppercase">Destination</label>
                        <input
                          type="text"
                          value={bulkDestination}
                          onChange={(e) => setBulkDestination(e.target.value)}
                          className="w-full bg-slate-950 border border-slate-800 text-slate-200 px-3 py-1.5 rounded text-xs font-semibold"
                        />
                      </div>

                      {/* Status */}
                      <div>
                        <label className="text-[10px] font-bold text-slate-400 uppercase">Status</label>
                        <select
                          value={bulkStatus}
                          onChange={(e) => setBulkStatus(e.target.value as any)}
                          className="w-full bg-slate-950 border border-slate-800 text-slate-200 px-2 py-1.5 rounded text-xs cursor-pointer font-semibold text-indigo-400"
                        >
                          <option value="Expected">Expected (Not Arrived)</option>
                          <option value="Received">Arrived (Received)</option>
                        </select>
                      </div>

                      {/* Damaged */}
                      <div>
                        <label className="text-[10px] font-bold text-slate-400 uppercase block mb-1">Damaged</label>
                        <div className="flex gap-2">
                          <button
                            type="button"
                            onClick={() => setBulkDamaged('Y')}
                            className={`flex-1 py-1 px-3 text-xs rounded font-semibold border transition cursor-pointer ${
                              bulkDamaged === 'Y'
                                ? 'bg-amber-600 border-amber-500 text-white shadow'
                                : 'bg-slate-950 border-slate-800 text-slate-400 hover:border-slate-700'
                            }`}
                          >
                            Y
                          </button>
                          <button
                            type="button"
                            onClick={() => setBulkDamaged('N')}
                            className={`flex-1 py-1 px-3 text-xs rounded font-semibold border transition cursor-pointer ${
                              bulkDamaged === 'N'
                                ? 'bg-indigo-600 border-indigo-500 text-white shadow'
                                : 'bg-slate-950 border-slate-800 text-slate-400 hover:border-slate-700'
                            }`}
                          >
                            N
                          </button>
                        </div>
                      </div>

                      {/* Locked (LN) */}
                      <div>
                        <label className="text-[10px] font-bold text-slate-400 uppercase block mb-1">Locked (L/N)</label>
                        <input
                          type="text"
                          placeholder="e.g. PAD, CL, Y, N"
                          value={bulkLn}
                          onChange={(e) => setBulkLn(e.target.value)}
                          className="w-full bg-slate-950 border border-slate-800 text-slate-200 px-3 py-1.5 rounded text-xs font-mono"
                        />
                      </div>

                      {/* Weight */}
                      <div>
                        <label className="text-[10px] font-bold text-slate-400 uppercase block mb-1">Weight (kg)</label>
                        <div className="relative">
                          <input
                            type="number"
                            step="0.1"
                            min="0"
                            placeholder="e.g. 23.4"
                            value={bulkWeight === undefined ? '' : bulkWeight}
                            onChange={(e) => setBulkWeight(e.target.value === '' ? undefined : Number(e.target.value))}
                            className="w-full bg-slate-950 border border-slate-800 text-slate-200 pl-3 pr-8 py-1.5 rounded text-xs"
                          />
                          <span className="absolute right-2.5 top-1/2 -translate-y-1/2 text-[10px] text-slate-500 font-bold">kg</span>
                        </div>
                      </div>

                      {/* Seal */}
                      <div className="col-span-2">
                        <label className="text-[10px] font-bold text-slate-400 uppercase">Seal</label>
                        <input
                          type="text"
                          placeholder="S-xxxxx"
                          value={bulkSeal}
                          onChange={(e) => setBulkSeal(e.target.value)}
                          className="w-full bg-slate-950 border border-slate-800 text-slate-200 px-2 py-1.5 rounded text-xs font-mono"
                        />
                      </div>
                    </div>

                    <button
                      type="submit"
                      className="w-full bg-gradient-to-r from-indigo-600 to-indigo-500 hover:from-indigo-500 hover:to-indigo-400 text-white font-bold py-3 rounded-lg text-xs cursor-pointer shadow-lg shadow-indigo-600/10 transition flex items-center justify-center gap-2"
                    >
                      <Check className="w-4 h-4" />
                      <span>Register and Save Bulk Operations ({parseBulkInput(bulkTagsInput).length} Records)</span>
                    </button>
                  </form>
                </div>
              )}
              </div>
            )}
          </section>
        )}

        {activeSection === 'protocol' && (
          <section id="protocol" className="max-w-4xl mx-auto bg-slate-900 border border-slate-800 rounded-xl p-6 space-y-4 w-full">
            <h3 className="font-bold text-slate-200 text-sm border-b border-slate-800 pb-2">Dispositions & Customs Operations Protocol</h3>
            
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div className="bg-slate-950 p-4 rounded-lg border border-slate-800">
                <h4 className="text-xs font-bold text-emerald-400 uppercase tracking-wide flex items-center gap-1.5 mb-2">
                  <span className="w-2 h-2 rounded-full bg-emerald-400" />
                  Cleared Baggage Options
                </h4>
                <ul className="text-xs text-slate-400 space-y-1.5 list-disc pl-4">
                  <li><strong>Delivery Agents:</strong> VVM, Outlook, or Advik.</li>
                  <li><strong>Storage:</strong> Standard warehousing at LHG Office.</li>
                  <li><strong>Domestic Forwarding:</strong> Forward via Air India, Indigo, or Spice Jet.</li>
                </ul>
              </div>

              <div className="bg-slate-950 p-4 rounded-lg border border-slate-800">
                <h4 className="text-xs font-bold text-red-400 uppercase tracking-wide flex items-center gap-1.5 mb-2">
                  <span className="w-2 h-2 rounded-full bg-red-400" />
                  Non-Cleared / Other Protocol
                </h4>
                <ul className="text-xs text-slate-400 space-y-1.5 list-disc pl-4">
                  <li><strong>Arrival Belt 9:</strong> Default holds and queue check.</li>
                  <li><strong>Handover:</strong> Handover custody to partner airlines.</li>
                  <li><strong>CWC Warehouse:</strong> Secure central depot storage.</li>
                  <li><strong>Re-export:</strong> Repatriate bag back to carrier hub.</li>
                </ul>
              </div>
            </div>

            <div className="bg-indigo-950/20 border border-indigo-900/30 p-4 rounded-lg text-xs text-indigo-300 flex items-start gap-2">
              <Clock className="w-4 h-4 shrink-0 mt-0.5" />
              <div>
                <p className="font-semibold">Simulating Days in Storage Reminder Alerts:</p>
                <p className="mt-1 text-slate-400">
                  Because new imports default to today&apos;s date, you can edit any bag and click the <strong className="text-indigo-400">&quot;Simulate 3 Days&quot;</strong> or <strong className="text-indigo-400">&quot;Simulate 5 Days&quot;</strong> shortcut buttons on received storage bags in the table below to backdate their entry dates instantly.
                </p>
              </div>
            </div>
          </section>
        )}

        {/* 6. Main Database & Multi-select list */}
        {activeSection === 'registry' && (
          <section className="bg-slate-900 border border-slate-800 rounded-xl overflow-hidden shadow-xl">
          
          {/* Header toolbar */}
          <div className="p-5 border-b border-slate-800 flex flex-col md:flex-row justify-between items-stretch md:items-center gap-4 bg-slate-900/80">
            <div className="flex flex-col sm:flex-row items-start sm:items-center gap-2">
              <h3 className="font-bold text-slate-100 text-sm">Main Operations Registry</h3>
              <div className="flex flex-wrap gap-1.5">
                <span className="bg-indigo-950/80 border border-indigo-800/40 text-indigo-400 text-[10px] font-mono px-2 py-0.5 rounded">
                  {summary.filteredRecords} Filtered / {summary.totalRecords} Total
                </span>
                <span className="bg-emerald-950/80 border border-emerald-800/40 text-emerald-400 text-[10px] font-mono px-2 py-0.5 rounded">
                  Arrivals: {summary.arrival}
                </span>
                <span className="bg-amber-950/80 border border-amber-800/40 text-amber-400 text-[10px] font-mono px-2 py-0.5 rounded">
                  Departures: {summary.departure}
                </span>
              </div>
            </div>

            <div className="flex flex-col sm:flex-row items-stretch sm:items-center gap-2">
              <div className="flex items-center gap-2">
                <button
                  onClick={exportToExcel}
                  className="flex-1 sm:flex-initial flex items-center justify-center gap-1.5 bg-slate-800 hover:bg-slate-700 text-slate-200 text-xs font-bold px-3 py-1.5 rounded cursor-pointer transition"
                  title="Download Excel Report"
                >
                  <Download className="w-3.5 h-3.5" />
                  Excel
                </button>
                
                {isMobile && (
                  <button
                    onClick={handleShareViaWhatsApp}
                    className="flex-1 sm:flex-initial flex items-center justify-center gap-1.5 bg-emerald-600/20 border border-emerald-500/20 hover:bg-emerald-600/30 text-emerald-400 text-xs font-bold px-3 py-1.5 rounded cursor-pointer transition"
                    title="Share via WhatsApp/Native"
                  >
                    <Share2 className="w-3.5 h-3.5" />
                    Share
                  </button>
                )}

                <button
                  onClick={handleShareViaEmail}
                  className="flex-1 sm:flex-initial flex items-center justify-center gap-1.5 bg-indigo-600/20 border border-indigo-500/20 hover:bg-indigo-600/30 text-indigo-400 text-xs font-bold px-3 py-1.5 rounded cursor-pointer transition"
                  title="Send via Email"
                >
                  <Mail className="w-3.5 h-3.5" />
                  Email
                </button>

                <button
                  onClick={copyToClipboardTable}
                  className="flex-1 sm:flex-initial flex items-center justify-center gap-1.5 bg-amber-600/20 border border-amber-500/20 hover:bg-amber-600/30 text-amber-400 text-xs font-bold px-3 py-1.5 rounded cursor-pointer transition"
                  title="Copy Table to Clipboard"
                >
                  <Copy className="w-3.5 h-3.5" />
                  Copy
                </button>
              </div>

              {/* Registry filter dropdown */}
              <div className="flex items-center gap-1.5">
                <Filter className="w-3.5 h-3.5 text-slate-400" />
                <select
                  value={filterRegistry}
                  onChange={(e) => setFilterRegistry(e.target.value as any)}
                  className="bg-slate-950 border border-slate-800 text-slate-300 text-xs rounded px-2.5 py-1.5 focus:border-indigo-500"
                >
                  <option value="Combined">Both Registries</option>
                  <option value="Arrival">Arrival Only</option>
                  <option value="Departure">Departure Only</option>
                </select>
              </div>

              {/* Date Filter */}
              <div className="flex items-center gap-1.5">
                <Calendar className="w-3.5 h-3.5 text-slate-400" />
                <select
                  value={filterDateRange.type}
                  onChange={(e) => {
                    const val = e.target.value as any;
                    if (val === 'custom') {
                      setFilterDateRange({ type: 'custom', from: '', to: '' });
                    } else {
                      setFilterDateRange({ type: val });
                    }
                  }}
                  className="bg-slate-950 border border-slate-800 text-slate-300 text-xs rounded px-2.5 py-1.5 focus:border-indigo-500"
                >
                  <option value="all">Any Date</option>
                  <option value="today">Today</option>
                  <option value="yesterday">Yesterday</option>
                  <option value="last7">Last 7 Days</option>
                  <option value="last30">Last 30 Days</option>
                  <option value="thisMonth">This Month</option>
                  <option value="prevMonth">Previous Month</option>
                  <option value="custom">Custom Range...</option>
                </select>
              </div>

              {filterDateRange.type === 'custom' && (
                <div className="flex items-center gap-1.5">
                  <input
                    type="date"
                    value={filterDateRange.from || ''}
                    onChange={(e) => setFilterDateRange(prev => ({ ...prev, from: e.target.value }))}
                    className="bg-slate-950 border border-slate-800 text-slate-300 text-[10px] rounded px-2 py-1 focus:border-indigo-500"
                  />
                  <span className="text-slate-500 text-[10px]">to</span>
                  <input
                    type="date"
                    value={filterDateRange.to || ''}
                    onChange={(e) => setFilterDateRange(prev => ({ ...prev, to: e.target.value }))}
                    className="bg-slate-950 border border-slate-800 text-slate-300 text-[10px] rounded px-2 py-1 focus:border-indigo-500"
                  />
                </div>
              )}

              {/* Flight filter dropdown */}
              <div className="flex items-center gap-1.5">
                <Filter className="w-3.5 h-3.5 text-slate-400" />
                <input
                  type="text"
                  placeholder="Flight(s) e.g. LH, LX154"
                  value={filterFlights.join(', ')}
                  onChange={(e) => {
                    const tokens = e.target.value.split(',').map(t => t.trim()).filter(t => t.length > 0);
                    setFilterFlights(tokens);
                  }}
                  className="bg-slate-950 border border-slate-800 text-slate-300 text-xs rounded px-2.5 py-1.5 focus:border-indigo-500 w-40"
                />
              </div>

              {/* Text search */}
              <div className="relative">
                <Search className="absolute left-2.5 top-1/2 -translate-y-1/2 w-3.5 h-3.5 text-slate-500" />
                <input
                  type="text"
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                  placeholder="Search Name, PIR, Tags, Remarks..."
                  className="bg-slate-950 border border-slate-800 pl-8 pr-3 py-1.5 rounded text-xs focus:border-indigo-500 text-slate-200 placeholder-slate-500 w-full sm:w-48"
                />
              </div>

              {/* Remarks filter */}
              <div className="flex items-center gap-1.5">
                <FileText className="w-3.5 h-3.5 text-slate-400" />
                <input
                  type="text"
                  placeholder="Filter by Remarks..."
                  value={filterRemarks}
                  onChange={(e) => setFilterRemarks(e.target.value)}
                  className="bg-slate-950 border border-slate-800 text-slate-300 text-xs rounded px-2.5 py-1.5 focus:border-indigo-500 w-40"
                />
              </div>

              {/* Clear filters shortcut */}
              {(searchTerm || filterFlights.length > 0 || filterRemarks || filterDateRange.type !== 'all' || activeFilter.type !== 'all') && (
                <button
                  onClick={handleClearFilters}
                  className="px-2.5 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded text-xs font-semibold flex items-center gap-1 transition"
                >
                  <RotateCcw className="w-3.5 h-3.5" />
                  <span className="hidden sm:inline">Clear</span>
                </button>
              )}
            </div>
          </div>

          {/* Bulk editing controls (Only displays when records are selected) */}
          {selectedIds.length > 0 && (
            <div id="bulk-controls" className="bg-indigo-950/30 border-b border-indigo-900/50 px-5 py-3 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
              <div className="flex items-center gap-2">
                <CheckCircle className="w-4 h-4 text-indigo-400" />
                <span className="text-xs text-indigo-300 font-semibold">
                  {selectedIds.length} records selected for Bulk Editing
                </span>
                <button
                  onClick={() => setSelectedIds([])}
                  className="text-[10px] text-slate-400 hover:text-slate-200 underline ml-2"
                >
                  Clear Selection
                </button>
              </div>

              <div className="flex flex-wrap items-center gap-2">
                <button
                  onClick={() => {
                    // Quick bulk receive
                    const updated = baggageList.map(item => {
                      if (selectedIds.includes(item.id)) {
                        return {
                          ...item,
                          status: 'Received' as const,
                          receivedAt: item.receivedAt || new Date().toISOString(),
                          disposition: item.disposition === 'Pending' ? ('Storage' as const) : item.disposition,
                          dispositionLocation: item.dispositionLocation || ('LHG Office' as const)
                        };
                      }
                      return item;
                    });
                    saveBaggageData(updated);
                    setSelectedIds([]);
                    alert("Selected bags marked as Received and routed to storage!");
                  }}
                  className="bg-emerald-600/20 hover:bg-emerald-600/30 text-emerald-300 border border-emerald-500/20 px-2.5 py-1 rounded text-xs transition"
                >
                  Bulk Mark Received
                </button>

                <button
                  onClick={() => {
                    // Open Bulk Edit Customs/Disposition Modal
                    setShowBulkEdit(true);
                  }}
                  className="bg-indigo-600 hover:bg-indigo-500 text-white px-2.5 py-1 rounded text-xs transition font-semibold"
                >
                  Edit Status & Disposition
                </button>

                <button
                  onClick={handleBulkDelete}
                  className="bg-red-600/20 hover:bg-red-600/30 text-red-300 border border-red-500/20 px-2.5 py-1 rounded text-xs transition"
                >
                  Delete ({selectedIds.length})
                </button>
              </div>
            </div>
          )}

          {/* Sub-tabs for specific workflows */}
          <div className="flex border-b border-slate-800 bg-slate-900/40">
            <button
              onClick={() => setActiveFilter({ type: 'all' })}
              className={`px-5 py-3 text-xs font-bold border-b-2 transition ${
                activeFilter.type === 'all' 
                  ? 'border-indigo-500 text-indigo-400' 
                  : 'border-transparent text-slate-400 hover:text-slate-200'
              }`}
            >
              All Bags
            </button>
            <button
              onClick={() => setActiveFilter({ type: 'non-arrival' })}
              className={`px-5 py-3 text-xs font-bold border-b-2 transition ${
                activeFilter.type === 'non-arrival' 
                  ? 'border-indigo-500 text-indigo-400' 
                  : 'border-transparent text-slate-400 hover:text-slate-200'
              }`}
            >
              Non-Arrivals Follow-up
            </button>
            <button
              onClick={() => {
                setActiveFilter({ type: 'location', value: 'LHG Office' });
              }}
              className={`px-5 py-3 text-xs font-bold border-b-2 transition cursor-pointer ${
                activeFilter.type === 'location' && activeFilter.value === 'LHG Office'
                  ? 'border-indigo-500 text-indigo-400' 
                  : 'border-transparent text-slate-400 hover:text-slate-200'
              }`}
            >
              Active Warehousing
            </button>
            <button
              onClick={() => {
                setActiveFilter({ type: 'alerts' });
              }}
              className={`px-5 py-3 text-xs font-bold border-b-2 transition cursor-pointer ${
                activeFilter.type === 'alerts' 
                  ? 'border-indigo-500 text-indigo-400' 
                  : 'border-transparent text-slate-400 hover:text-slate-200'
              }`}
            >
              Storage Alerts Tracker ({activeAlerts.totalAlerts})
            </button>
          </div>

          {/* Database Table Container */}
          <div className="overflow-x-auto">
            {filteredBaggage.length === 0 ? (
              <div className="p-8 text-center text-slate-500 space-y-2">
                <p>No baggage records matched the selected query, filters, or status.</p>
                <button
                  onClick={handleClearFilters}
                  className="text-xs text-indigo-400 hover:underline font-semibold"
                >
                  Reset Dashboard Filters & Search
                </button>
              </div>
            ) : (
              <table className="w-full text-left border-collapse text-xs">
                <thead>
                  <tr className="bg-slate-950/80 text-slate-400 uppercase tracking-wider font-mono border-b border-slate-800">
                    <th className="py-3 px-4 w-10">
                      <input
                        type="checkbox"
                        checked={
                          filteredBaggage.length > 0 &&
                          filteredBaggage.every(r => selectedIds.includes(r.id))
                        }
                        onChange={() => handleSelectAll(filteredBaggage)}
                        className="rounded border-slate-800 text-indigo-600 focus:ring-indigo-500"
                      />
                    </th>
                    <th className="py-3 px-3 font-semibold text-slate-300">Flight No</th>
                    <th className="py-3 px-3 font-semibold text-slate-300">Receipt Date</th>
                    <th className="py-3 px-3">Original Tag</th>
                    <th className="py-3 px-3">Rush Tag</th>
                    <th className="py-3 px-3">Passenger Name</th>
                    <th className="py-3 px-3">PIR Number</th>
                    <th className="py-3 px-3">Operational Specs</th>
                    <th className="py-3 px-3">Customs Status</th>
                    <th className="py-3 px-3">Disposition & Storage Location</th>
                    <th className="py-3 px-3 text-center">Age in Store</th>
                    <th className="py-3 px-3 text-right">Actions</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-800/80">
                  {filteredBaggage.map((record) => {
                    const daysInStorage = getDaysInStorage(record);
                    const is3DayAlert = daysInStorage >= 3 && daysInStorage < 5;
                    const is5DayAlert = daysInStorage >= 5;
                    const isSelected = selectedIds.includes(record.id);

                    return (
                      <tr 
                        key={record.id}
                        onDoubleClick={() => handleOpenEditDialog(record)}
                        title="Double click to edit record"
                        className={`hover:bg-slate-900/60 cursor-pointer select-none transition ${
                          isSelected ? 'bg-indigo-950/20' : ''
                        }`}
                      >
                        {/* Selector checkbox */}
                        <td className="py-3 px-4">
                          <input
                            type="checkbox"
                            checked={isSelected}
                            onChange={() => handleSelectToggle(record.id)}
                            className="rounded border-slate-800 text-indigo-600 focus:ring-indigo-500"
                          />
                        </td>

                        {/* 1. Flight No */}
                        <td className="py-3 px-3">
                          <span className="px-2 py-0.5 bg-slate-950 rounded text-indigo-300 border border-slate-800 font-mono font-bold">
                            {record.flightNo}
                          </span>
                        </td>

                        {/* Date of Receipt */}
                        <td className="py-3 px-3">
                          <div className="text-[11px] font-bold text-slate-100 font-mono">
                            {record.receivedAt ? new Date(record.receivedAt).toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }).toUpperCase().replace(/ /g, '-') : '-'}
                          </div>
                          <div className="text-[9px] text-slate-500 font-mono">
                            {record.receivedAt ? new Date(record.receivedAt).toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit', hour12: false }) : ''}
                          </div>
                        </td>

                        {/* 2. Original Tag */}
                        <td className="py-3 px-3 font-mono text-slate-300">
                          {record.originalTag || <span className="text-slate-600">-</span>}
                        </td>

                        {/* 3. Rush Tag */}
                        <td className="py-3 px-3 font-mono">
                          {record.rushTag ? (
                            <span className={record.rushTag === 'Not Loaded' ? 'text-red-500 font-bold' : 'text-indigo-400 font-bold'}>
                              {record.rushTag === 'Not Loaded' ? record.rushTag : `Rush: ${record.rushTag}`}
                            </span>
                          ) : (
                            <span className="text-slate-600">-</span>
                          )}
                        </td>

                        {/* 4. Passenger Name */}
                        <td className="py-3 px-3 font-semibold text-slate-200">
                          {record.name}
                        </td>

                        {/* 5. PIR Number */}
                        <td className="py-3 px-3 font-mono text-slate-100 font-medium">
                          {record.pir || <span className="text-slate-600">NO PIR</span>}
                        </td>

                        {/* 6. Operational Specs (SNO, Locked, Weight, Damaged, Seal) */}
                        <td className="py-3 px-3 space-y-0.5">
                          <div className="text-[10px] text-slate-400 font-medium">
                            Locked: <span className={
                              /^(n|no|false)$/i.test(record.ln || '') || !record.ln
                                ? 'text-emerald-400 font-bold'
                                : 'text-red-400 font-bold'
                            }>{record.ln || 'No'}</span>
                            {record.sno && <span className="text-slate-500 ml-1.5">(SNo: {record.sno})</span>}
                          </div>
                          <div className="text-[10px] text-slate-300 font-semibold font-mono flex items-center gap-1.5">
                            {record.weight !== undefined ? (
                              <span className="bg-slate-950 px-1.5 py-0.5 rounded text-indigo-300">
                                {record.weight.toFixed(1)} kg
                              </span>
                            ) : null}
                            {record.damaged === 'Y' ? (
                              <span className="bg-red-950/60 border border-red-900 text-red-400 text-[9px] px-1 rounded font-bold">
                                DAMAGED
                              </span>
                            ) : null}
                          </div>
                          {record.seal && (
                            <div className="text-[9px] text-slate-500 font-mono">Seal: {record.seal}</div>
                          )}
                        </td>

                        {/* Customs Status */}
                        <td className="py-3 px-3">
                          {record.status === 'Expected' ? (
                            <span className="text-slate-500 italic">Expected Manifest</span>
                          ) : (
                            <div className="space-y-1">
                              {record.customsStatus === 'Cleared' ? (
                                <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full bg-emerald-950 text-emerald-400 border border-emerald-800 font-medium text-[10px]">
                                  <Check className="w-2.5 h-2.5" />
                                  Cleared
                                </span>
                              ) : record.customsStatus === 'Marked Preventive' ? (
                                <div className="space-y-0.5">
                                  <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full bg-red-950 text-amber-400 border border-amber-600/50 font-bold text-[10px] animate-pulse">
                                    <AlertOctagon className="w-2.5 h-2.5 text-amber-400" />
                                    Preventive Hold
                                  </span>
                                  <div className="text-[9px] text-amber-300 font-semibold leading-tight max-w-[150px]">
                                    Req: Passenger Arrival or HUB Re-export
                                  </div>
                                </div>
                              ) : record.customsStatus === 'Not Cleared' ? (
                                <div className="space-y-0.5">
                                  <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full bg-red-950 text-red-400 border border-red-800 font-medium text-[10px]">
                                    <AlertTriangle className="w-2.5 h-2.5" />
                                    Held Custom
                                  </span>
                                  {record.customsReason && (
                                    <div className="text-[9px] text-red-300 italic">({record.customsReason})</div>
                                  )}
                                </div>
                              ) : (
                                <span className="px-2 py-0.5 rounded-full bg-slate-800 text-slate-400 text-[10px]">
                                  Pending
                                </span>
                              )}
                            </div>
                          )}
                        </td>

                        {/* Disposition and location */}
                        <td className="py-3 px-3">
                          {record.status === 'Expected' ? (
                            <div className="flex items-center gap-1.5 text-slate-500">
                              <Clock className="w-3.5 h-3.5" />
                              <span>Not Received Yet</span>
                            </div>
                          ) : (
                            <div className="space-y-1">
                              <div className="font-semibold text-slate-200">
                                {record.dispositionLocation || 'No active placement'}
                              </div>
                              <div className="text-[10px] text-slate-400 font-mono capitalize">
                                Action: {record.disposition}
                              </div>
                            </div>
                          )}
                        </td>

                        {/* Storage Alert Indicator */}
                        <td className="py-3 px-3 text-center">
                          {daysInStorage > 0 ? (
                            <div className="flex flex-col items-center justify-center">
                              <span className={`px-2 py-0.5 rounded text-[11px] font-mono font-bold ${
                                is5DayAlert 
                                  ? 'bg-red-500 text-white animate-pulse'
                                  : is3DayAlert 
                                  ? 'bg-amber-500 text-slate-950'
                                  : 'bg-slate-800 text-slate-300'
                              }`}>
                                {daysInStorage} Days
                              </span>
                              {is5DayAlert && (
                                <span className="text-[9px] text-red-400 font-bold uppercase tracking-wider mt-0.5">RE-EXPORT</span>
                              )}
                              {is3DayAlert && (
                                <span className="text-[9px] text-amber-400 font-semibold uppercase tracking-wider mt-0.5">WARN</span>
                              )}
                            </div>
                          ) : record.status === 'Received' && record.dispositionLocation && ['LHG Office', 'Belt 9', 'BMA', 'Level 4 Checks', 'CWC'].includes(record.dispositionLocation) ? (
                            <span className="text-slate-400 font-mono">0 Days</span>
                          ) : (
                            <span className="text-slate-500">-</span>
                          )}
                        </td>

                        {/* Actions */}
                        <td className="py-3 px-3 text-right space-y-1">
                          <div className="flex justify-end gap-1.5">
                            {/* Receive shortcut button */}
                            {record.status === 'Expected' && (
                              <button
                                onClick={() => {
                                  const updated = baggageList.map(item => {
                                    if (item.id === record.id) {
                                      return {
                                        ...item,
                                        status: 'Received' as const,
                                        receivedAt: new Date().toISOString(),
                                        disposition: 'Storage' as const,
                                        dispositionLocation: 'LHG Office' as const,
                                        dispositionUpdatedAt: new Date().toISOString()
                                      };
                                    }
                                    return item;
                                  });
                                  saveBaggageData(updated);
                                }}
                                className="px-2 py-1 bg-emerald-600 hover:bg-emerald-500 text-white rounded text-[10px] font-semibold transition"
                              >
                                Mark Received
                              </button>
                            )}

                            {/* Edit Single Button */}
                            <button
                              onClick={() => handleOpenEditDialog(record)}
                              className="p-1 text-slate-400 hover:text-white bg-slate-800 hover:bg-slate-700 rounded transition"
                              title="Edit Record"
                            >
                              <Edit className="w-3.5 h-3.5" />
                            </button>

                            {/* Delete single button (Admin only checks inside handler) */}
                            <button
                              onClick={() => handleDeleteSingle(record.id)}
                              className="p-1 text-red-400 hover:text-red-300 bg-red-950/20 hover:bg-red-950/40 rounded transition"
                              title="Delete Record"
                            >
                              <Trash2 className="w-3.5 h-3.5" />
                            </button>
                          </div>

                          {/* Quick Backdate Alert Testing simulation tools */}
                          {record.status === 'Received' && (
                            <div className="flex justify-end gap-1 text-[9px]">
                              <span className="text-slate-500 self-center mr-1">Simulate store alert:</span>
                              <button
                                onClick={() => forceBackdateForAlert(record.id, 3)}
                                className="px-1.5 py-0.5 bg-amber-500/10 hover:bg-amber-500/20 text-amber-400 rounded transition border border-amber-500/20"
                              >
                                3 Days
                              </button>
                              <button
                                onClick={() => forceBackdateForAlert(record.id, 5)}
                                className="px-1.5 py-0.5 bg-red-500/10 hover:bg-red-500/20 text-red-400 rounded transition border border-red-500/20"
                              >
                                5 Days
                              </button>
                            </div>
                          )}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            )}
          </div>
        </section>
        )}
      </main>
    </div>

      {/* 7. Dialog Modals */}

      {/* A. EXCEL / BDO PASSED TEXT IMPORT DIALOG */}
      {showImportDialog && (
        <div className="fixed inset-0 bg-black/80 backdrop-blur-md z-50 flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-800 w-full max-w-5xl rounded-2xl shadow-2xl overflow-hidden flex flex-col max-h-[92vh] animate-in fade-in duration-200">
            
            {/* Header */}
            <div className="p-5 border-b border-slate-800 flex justify-between items-center bg-slate-900">
              <div className="flex items-center gap-3">
                <div className="p-2 bg-indigo-500/10 rounded-lg">
                  <FileSpreadsheet className="w-5 h-5 text-indigo-400" />
                </div>
                <div>
                  <h3 className="font-bold text-slate-100 text-sm flex items-center gap-2">
                    Intelligent Baggage Manifest Importer
                    <span className="text-[10px] bg-indigo-500/20 text-indigo-300 font-mono px-2 py-0.5 rounded font-normal uppercase tracking-wider">
                      Fuzzy + Semantic
                    </span>
                  </h3>
                  <p className="text-[11px] text-slate-400">Robust Excel / CSV / Pasted text reconciliation engine</p>
                </div>
              </div>
              <div className="flex items-center gap-2">
                <button 
                  onClick={() => {
                    setShowDictionaryEditor(!showDictionaryEditor);
                    setShowLockDictionaryEditor(false);
                  }}
                  className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold transition ${
                    showDictionaryEditor 
                      ? 'bg-amber-500/20 text-amber-300 border border-amber-500/30' 
                      : 'bg-slate-800 hover:bg-slate-700 text-slate-300 border border-slate-700'
                  }`}
                  title="Configure fuzzy synonyms & aliases"
                >
                  <Settings className="w-3.5 h-3.5 text-amber-400" />
                  Mapping Dictionary
                </button>
                <button 
                  onClick={() => {
                    setShowLockDictionaryEditor(!showLockDictionaryEditor);
                    setShowDictionaryEditor(false);
                  }}
                  className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold transition ${
                    showLockDictionaryEditor 
                      ? 'bg-indigo-500/20 text-indigo-300 border border-indigo-500/30' 
                      : 'bg-slate-800 hover:bg-slate-700 text-slate-300 border border-slate-700'
                  }`}
                  title="Configure Lock Abbreviation Dictionary"
                >
                  <Lock className="w-3.5 h-3.5 text-indigo-400" />
                  Lock Dictionary
                </button>
                <button 
                  onClick={() => {
                    setShowImportDialog(false);
                    setImportWizardStep('upload');
                    setExcelRows([]);
                    setExcelHeaders([]);
                    setRawPasteText('');
                    setImportSummaryResult(null);
                    setIsSBHImport(false);
                    setImportProgressMessage('');
                  }}
                  className="p-1.5 bg-slate-800/80 hover:bg-slate-800 rounded-lg text-slate-400 hover:text-slate-100 transition"
                >
                  <X className="w-4 h-4" />
                </button>
              </div>
            </div>

            {/* Step Wizard Progress Bar */}
            <div className="bg-slate-950 px-6 py-3 border-b border-slate-800/50 flex justify-between items-center text-xs font-mono">
              <div className="flex items-center gap-6">
                <button
                  disabled={importWizardStep === 'summary'}
                  onClick={() => setImportWizardStep('upload')}
                  className={`flex items-center gap-1.5 transition ${
                    importWizardStep === 'upload' ? 'text-indigo-400 font-bold' : 'text-slate-500 hover:text-slate-300'
                  }`}
                >
                  <span className={`w-5 h-5 rounded-full flex items-center justify-center text-[10px] ${
                    importWizardStep === 'upload' ? 'bg-indigo-500 text-slate-950 font-bold' : 'bg-slate-800'
                  }`}>1</span>
                  Source
                </button>
                <span className="text-slate-700">/</span>
                <button
                  disabled={excelHeaders.length === 0 || importWizardStep === 'summary'}
                  onClick={() => setImportWizardStep('mapping')}
                  className={`flex items-center gap-1.5 transition ${
                    importWizardStep === 'mapping' ? 'text-indigo-400 font-bold' : 'text-slate-500 hover:text-slate-300'
                  }`}
                >
                  <span className={`w-5 h-5 rounded-full flex items-center justify-center text-[10px] ${
                    importWizardStep === 'mapping' ? 'bg-indigo-500 text-slate-950 font-bold' : 'bg-slate-800'
                  }`}>2</span>
                  Column Mapping
                </button>
                <span className="text-slate-700">/</span>
                <button
                  disabled={excelHeaders.length === 0 || importWizardStep === 'summary'}
                  onClick={() => setImportWizardStep('preview')}
                  className={`flex items-center gap-1.5 transition ${
                    importWizardStep === 'preview' ? 'text-indigo-400 font-bold' : 'text-slate-500 hover:text-slate-300'
                  }`}
                >
                  <span className={`w-5 h-5 rounded-full flex items-center justify-center text-[10px] ${
                    importWizardStep === 'preview' ? 'bg-indigo-500 text-slate-950 font-bold' : 'bg-slate-800'
                  }`}>3</span>
                  Preview Rows
                </button>
                <span className="text-slate-700">/</span>
                <div className={`flex items-center gap-1.5 ${
                  importWizardStep === 'summary' ? 'text-indigo-400 font-bold' : 'text-slate-500'
                }`}>
                  <span className={`w-5 h-5 rounded-full flex items-center justify-center text-[10px] ${
                    importWizardStep === 'summary' ? 'bg-indigo-500 text-slate-950 font-bold' : 'bg-slate-800'
                  }`}>4</span>
                  Import Log
                </div>
              </div>
              <div>
                {importProgress !== null && (
                  <div className="flex flex-col items-end">
                    <div className="flex items-center gap-2">
                      <RefreshCw className="w-3 h-3 text-indigo-400 animate-spin" />
                      <span className="text-indigo-300">Working {importProgress}%</span>
                    </div>
                    {importProgressMessage && (
                      <span className="text-[9px] text-slate-500 font-mono mt-0.5">{importProgressMessage}</span>
                    )}
                  </div>
                )}
              </div>
            </div>

            {/* Dictionary Editor Panel (Slide down overlay) */}
            {showDictionaryEditor && (
              <div className="bg-slate-950 border-b border-slate-800 p-5 space-y-4 max-h-[350px] overflow-y-auto">
                <div className="flex justify-between items-center">
                  <div>
                    <h4 className="text-xs font-bold text-slate-200 uppercase tracking-wider flex items-center gap-1.5">
                      <Sparkles className="w-3.5 h-3.5 text-amber-400" />
                      Configure Extensible Mapping Dictionary
                    </h4>
                    <p className="text-[10px] text-slate-400 mt-0.5">Add custom synonyms/aliases to let the system auto-recognize your Excel headers instantly.</p>
                  </div>
                  <button 
                    onClick={() => {
                      if (confirm('Reset custom mapping dictionary to factory defaults?')) {
                        saveMappingDictionary(DEFAULT_MAPPING_DICTIONARY);
                      }
                    }}
                    className="text-[10px] text-slate-400 hover:text-slate-200 underline font-mono flex items-center gap-1"
                  >
                    <RotateCcw className="w-3 h-3" />
                    Reset to Defaults
                  </button>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3">
                  {mappingDictionary.map((entry) => (
                    <div key={entry.field} className="bg-slate-900 p-3 rounded-lg border border-slate-800 space-y-2">
                      <div className="flex justify-between items-start">
                        <div>
                          <span className="text-xs font-bold text-slate-200">{entry.label}</span>
                          <span className="text-[9px] text-slate-500 block font-mono">{entry.field} {entry.isMandatory ? '(Mandatory)' : '(Optional)'}</span>
                        </div>
                      </div>
                      
                      <div className="flex flex-wrap gap-1">
                        {entry.aliases.map((alias) => (
                          <span 
                            key={alias} 
                            className="inline-flex items-center gap-1 text-[9px] bg-slate-950 text-indigo-300 border border-slate-800 px-1.5 py-0.5 rounded font-mono"
                          >
                            {alias}
                            <button 
                              onClick={() => {
                                const updated = mappingDictionary.map(x => {
                                  if (x.field === entry.field) {
                                    return { ...x, aliases: x.aliases.filter(a => a !== alias) };
                                  }
                                  return x;
                                });
                                saveMappingDictionary(updated);
                              }}
                              className="text-slate-500 hover:text-red-400 ml-1 font-bold"
                            >
                              ×
                            </button>
                          </span>
                        ))}
                      </div>
                    </div>
                  ))}
                </div>

                {/* Add Alias Form */}
                <div className="bg-slate-900/60 p-3 rounded-lg border border-slate-800/80 flex flex-wrap gap-2 items-center">
                  <span className="text-xs text-slate-400">Quick-Add Alias:</span>
                  <select 
                    value={newAliasField} 
                    onChange={(e) => setNewAliasField(e.target.value)}
                    className="bg-slate-950 border border-slate-800 rounded px-2 py-1 text-xs text-slate-300"
                  >
                    <option value="">-- Choose Field --</option>
                    {mappingDictionary.map(e => (
                      <option key={e.field} value={e.field}>{e.label}</option>
                    ))}
                  </select>
                  <input 
                    type="text" 
                    placeholder="e.g. flight_id, flt#, wt, pax_nm"
                    value={newAliasValue}
                    onChange={(e) => setNewAliasValue(e.target.value)}
                    className="bg-slate-950 border border-slate-800 rounded px-3 py-1 text-xs text-slate-200 flex-1 min-w-[200px]"
                  />
                  <button 
                    onClick={() => {
                      if (!newAliasField || !newAliasValue.trim()) return;
                      const val = newAliasValue.toLowerCase().trim();
                      const updated = mappingDictionary.map(x => {
                        if (x.field === newAliasField) {
                          if (x.aliases.includes(val)) return x;
                          return { ...x, aliases: [...x.aliases, val] };
                        }
                        return x;
                      });
                      saveMappingDictionary(updated);
                      setNewAliasValue('');
                    }}
                    className="bg-indigo-600 hover:bg-indigo-500 text-white font-bold px-4 py-1 rounded text-xs transition"
                  >
                    Add Alias
                  </button>
                </div>
              </div>
            )}

            {/* Lock Dictionary Editor Panel (Slide down overlay) */}
            {showLockDictionaryEditor && (
              <div className="bg-slate-950 border-b border-slate-800 p-5 space-y-4 max-h-[350px] overflow-y-auto animate-in slide-in-from-top duration-200">
                <div className="flex justify-between items-center">
                  <div>
                    <h4 className="text-xs font-bold text-slate-200 uppercase tracking-wider flex items-center gap-1.5 font-sans">
                      <Lock className="w-3.5 h-3.5 text-indigo-400" />
                      Configure Configurable Lock Mapping Dictionary
                    </h4>
                    <p className="text-[10px] text-slate-400 mt-0.5">Define custom lock abbreviations and how they automatically expand on import (e.g., CL to Combination Lock).</p>
                  </div>
                  <button 
                    onClick={() => {
                      if (confirm('Reset custom lock dictionary to factory defaults?')) {
                        saveLockDictionary(DEFAULT_LOCK_DICTIONARY);
                      }
                    }}
                    className="text-[10px] text-slate-400 hover:text-slate-200 underline font-mono flex items-center gap-1"
                  >
                    <RotateCcw className="w-3 h-3" />
                    Reset to Defaults
                  </button>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3">
                  {Object.entries(lockDictionary).map(([abbr, expanded]) => (
                    <div key={abbr} className="bg-slate-900 p-3 rounded-lg border border-slate-800 flex justify-between items-center">
                      <div>
                        <span className="text-xs font-bold text-indigo-300 font-mono bg-indigo-950/30 px-1.5 py-0.5 rounded border border-indigo-900/30">{abbr}</span>
                        <span className="text-[10px] text-slate-300 block mt-2 font-medium">{expanded}</span>
                      </div>
                      <button 
                        onClick={() => {
                          const updated = { ...lockDictionary };
                          delete updated[abbr];
                          saveLockDictionary(updated);
                        }}
                        className="text-slate-500 hover:text-red-400 font-bold p-1 bg-slate-950 rounded border border-slate-800 hover:border-red-900/40 text-xs w-6 h-6 flex items-center justify-center transition"
                        title="Remove Mapping"
                      >
                        ×
                      </button>
                    </div>
                  ))}
                </div>

                {/* Add Lock Mapping Form */}
                <div className="bg-slate-900/60 p-3 rounded-lg border border-slate-800/80 flex flex-wrap gap-2 items-center">
                  <span className="text-xs text-slate-400 font-semibold">Add Lock Mapping:</span>
                  <input 
                    type="text" 
                    placeholder="Abbreviation (e.g. CL)"
                    value={newLockAbbr}
                    onChange={(e) => setNewLockAbbr(e.target.value)}
                    className="bg-slate-950 border border-slate-800 rounded px-2.5 py-1 text-xs text-slate-200 w-44 font-mono uppercase focus:border-indigo-500 focus:outline-none"
                  />
                  <input 
                    type="text" 
                    placeholder="Expanded description (e.g. Combination Lock)"
                    value={newLockExpanded}
                    onChange={(e) => setNewLockExpanded(e.target.value)}
                    className="bg-slate-950 border border-slate-800 rounded px-3 py-1 text-xs text-slate-200 flex-1 min-w-[200px] focus:border-indigo-500 focus:outline-none"
                  />
                  <button 
                    onClick={() => {
                      if (!newLockAbbr.trim() || !newLockExpanded.trim()) return;
                      const key = newLockAbbr.toUpperCase().trim();
                      const val = newLockExpanded.trim();
                      const updated = { ...lockDictionary, [key]: val };
                      saveLockDictionary(updated);
                      setNewLockAbbr('');
                      setNewLockExpanded('');
                    }}
                    className="bg-indigo-600 hover:bg-indigo-500 text-white font-bold px-4 py-1 rounded text-xs transition"
                  >
                    Add Mapping
                  </button>
                </div>
              </div>
            )}

            {/* Main Content Scroll Area */}
            <div className="p-6 overflow-y-auto flex-1 bg-slate-900/40">
              
              {/* STEP 1: UPLOAD / RAW PASTE SOURCE */}
              {importWizardStep === 'upload' && (
                <div className="space-y-6">
                  <div className={`grid grid-cols-1 ${isMobile ? 'md:grid-cols-3' : 'md:grid-cols-2'} gap-4 text-xs`}>
                    <button
                      onClick={() => setImportTab('paste')}
                      className={`p-4 rounded-xl border text-left transition relative ${
                        importTab === 'paste' 
                          ? 'border-indigo-500 bg-indigo-950/20 text-indigo-200' 
                          : 'border-slate-800 bg-slate-950/60 text-slate-400 hover:text-slate-200'
                      }`}
                    >
                      <div className="flex items-center gap-2 mb-1.5">
                        <FileText className="w-4 h-4 text-indigo-400" />
                        <span className="font-bold text-slate-200">Option 1: Direct Spreadsheet Paste</span>
                      </div>
                      <p className="text-[10px] text-slate-400 leading-relaxed">
                        Copy columns directly from your active Excel sheet, Air India, or Lufthansa portal and paste them below. Zero file-saving required.
                      </p>
                    </button>

                    <button
                      onClick={() => setImportTab('file')}
                      className={`p-4 rounded-xl border text-left transition relative ${
                        importTab === 'file' 
                          ? 'border-indigo-500 bg-indigo-950/20 text-indigo-200' 
                          : 'border-slate-800 bg-slate-950/60 text-slate-400 hover:text-slate-200'
                      }`}
                    >
                      <div className="flex items-center gap-2 mb-1.5">
                        <Upload className="w-4 h-4 text-indigo-400" />
                        <span className="font-bold text-slate-200">Option 2: Import Spreadsheet Document</span>
                      </div>
                      <p className="text-[10px] text-slate-400 leading-relaxed">
                        Directly process your BDO file. Drag and drop any Excel file (.xlsx, .xls), standard CSV, or tab-delimited text sheets.
                      </p>
                    </button>

                    {isMobile && (
                      <button
                        onClick={() => {
                          setShowImportDialog(false);
                          setOcrTargetMode('table');
                          cameraInputRef.current?.click();
                        }}
                        className="p-4 rounded-xl border border-indigo-500/30 bg-indigo-950/20 text-indigo-200 text-left transition relative"
                      >
                        <div className="flex items-center gap-2 mb-1.5">
                          <Camera className="w-4 h-4 text-indigo-400" />
                          <span className="font-bold text-slate-200">Option 3: Camera Capture</span>
                        </div>
                        <p className="text-[10px] text-slate-400 leading-relaxed">
                          Photograph a printed baggage list or screen. Gemini AI will convert the image into structured data.
                        </p>
                      </button>
                    )}
                  </div>

                  {/* Duplicate Resolution Setting */}
                  <div className="bg-slate-950 p-4 rounded-xl border border-slate-800 flex justify-between items-center flex-wrap gap-4 text-xs">
                    <div>
                      <h4 className="font-bold text-slate-200">System Duplicate Resolution Protocol</h4>
                      <p className="text-[10px] text-slate-500">How would you like Swiss Terminal to treat rows matching existing PIR or tags?</p>
                    </div>
                    <div className="flex items-center gap-2">
                      <button 
                        onClick={() => setDuplicateMode('skip')}
                        className={`px-3 py-1.5 rounded-lg border font-semibold font-mono text-[11px] transition ${
                          duplicateMode === 'skip' 
                            ? 'bg-indigo-600 border-indigo-500 text-white' 
                            : 'bg-slate-900 border-slate-800 text-slate-400 hover:text-slate-200'
                        }`}
                      >
                        Skip Duplicates
                      </button>
                      <button 
                        onClick={() => setDuplicateMode('update')}
                        className={`px-3 py-1.5 rounded-lg border font-semibold font-mono text-[11px] transition ${
                          duplicateMode === 'update' 
                            ? 'bg-amber-600 border-amber-500 text-white' 
                            : 'bg-slate-900 border-slate-800 text-slate-400 hover:text-slate-200'
                        }`}
                      >
                        Overwrite & Update
                      </button>
                    </div>
                  </div>

                  {importTab === 'paste' ? (
                    <div className="space-y-3">
                      <div className="flex justify-between items-center">
                        <label className="text-[11px] font-bold text-slate-300 uppercase tracking-wide">Excel Clipboard Buffer</label>
                        <span className="text-[10px] text-slate-500 italic">Tabs, spaces, and CSV formats automatically aligned</span>
                      </div>
                      <textarea
                        rows={8}
                        value={rawPasteText}
                        onChange={(e) => setRawPasteText(e.target.value)}
                        placeholder="Paste Excel lines here (including header labels or raw columns)..."
                        className="w-full bg-slate-950 border border-slate-800 focus:border-indigo-500 text-slate-200 p-4 rounded-xl text-xs font-mono focus:ring-1 focus:ring-indigo-500/20 outline-none"
                      />
                      <div className="flex justify-between items-center">
                        <p className="text-[10px] text-slate-500">
                          Supports custom Column order! Row alignment occurs dynamically during parsing.
                        </p>
                        <div className="flex gap-2">
                          <button
                            onClick={() => setRawPasteText('')}
                            className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg text-xs font-semibold transition"
                          >
                            Reset
                          </button>
                          <button
                            onClick={parsePastedData}
                            disabled={!rawPasteText.trim() || isAnalyzing}
                            className="px-5 py-2 bg-indigo-600 hover:bg-indigo-500 disabled:opacity-50 text-white rounded-lg text-xs transition font-semibold flex items-center gap-1.5"
                          >
                            {isAnalyzing && <RefreshCw className="w-3.5 h-3.5 animate-spin" />}
                            Process Clipboard Data
                          </button>
                        </div>
                      </div>
                    </div>
                  ) : (
                    <div className="space-y-4">
                      <div className="border-2 border-dashed border-slate-800 rounded-xl p-12 text-center bg-slate-950/30 hover:bg-slate-950/60 transition relative">
                        <input
                          type="file"
                          accept=".xlsx,.xls,.csv,.txt"
                          multiple
                          onChange={handleCSVUpload}
                          className="absolute inset-0 opacity-0 cursor-pointer"
                        />
                        <Upload className="w-10 h-10 text-indigo-400 mx-auto mb-3" />
                        <p className="text-xs font-bold text-slate-200">Drag & Drop or Click to select BDO Sheet</p>
                        <p className="text-[10px] text-slate-500 mt-1 max-w-md mx-auto">
                          Supports native Excel Spreadsheet (.xlsx, .xls), Standard UTF-8 CSV, and TXT Tab Delimited formats
                        </p>
                      </div>
                    </div>
                  )}
                </div>
              )}

              {/* STEP 2: COLUMN MATCHING & MAPPING DICTIONARY HIGHLIGHTS */}
              {importWizardStep === 'mapping' && (
                <div className="space-y-6">
                  <div className="flex justify-between items-start flex-wrap gap-4">
                    <div>
                      <h4 className="font-bold text-slate-200 text-sm">System Field - Column Matcher</h4>
                      <p className="text-xs text-slate-400 mt-0.5">
                        Please review the automatically recognized columns. Adjust any fields with low confidence using the dropdown menus.
                      </p>
                    </div>
                    <div className="flex items-center gap-2">
                      <button
                        onClick={() => setImportWizardStep('upload')}
                        className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg text-xs font-semibold"
                      >
                        Back
                      </button>
                      <button
                        onClick={() => setImportWizardStep('preview')}
                        className="px-5 py-2 bg-indigo-600 hover:bg-indigo-500 text-white font-bold rounded-lg text-xs transition flex items-center gap-1"
                      >
                        Next: Review Rows
                      </button>
                    </div>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    {excelHeaders.map((header, colIdx) => {
                      const currentMapping = columnMappings[colIdx] || { systemField: 'ignore', confidence: 0, matchedBy: 'manual' };
                      
                      // Highlight color classes based on confidence
                      const confidenceColorClass = currentMapping.confidence >= 80 
                        ? 'bg-emerald-950 text-emerald-400 border border-emerald-800/60' 
                        : currentMapping.confidence >= 40 
                          ? 'bg-amber-950 text-amber-400 border border-amber-800/60'
                          : currentMapping.systemField === 'ignore'
                            ? 'bg-slate-900 text-slate-500 border border-slate-800'
                            : 'bg-red-950 text-red-400 border border-red-800/60';

                      // Extract 3 data sample values for previewing
                      const samples = excelRows.slice(0, 3).map(row => {
                        const val = row[colIdx];
                        return val !== undefined && val !== null ? String(val) : '';
                      }).filter(Boolean);

                      return (
                        <div key={colIdx} className="bg-slate-950 p-4 rounded-xl border border-slate-800 flex flex-col justify-between gap-4">
                          <div className="flex justify-between items-start gap-4">
                            <div>
                              <span className="text-[10px] text-slate-500 font-mono block">EXCEL HEADER COLUMN {colIdx + 1}</span>
                              <h5 className="font-bold text-slate-200 mt-0.5 font-mono">&quot;{header}&quot;</h5>
                            </div>
                            <span className={`px-2 py-0.5 rounded text-[10px] font-mono font-bold ${confidenceColorClass}`}>
                              {currentMapping.systemField === 'ignore' ? 'Ignored' : `${currentMapping.confidence}% Match`}
                            </span>
                          </div>

                          {/* Data Samples */}
                          <div className="bg-slate-900/60 p-2 rounded border border-slate-800/40 text-[10px] font-mono text-slate-400 space-y-0.5">
                            <span className="text-[8px] text-slate-600 block uppercase">Data Samples:</span>
                            {samples.length === 0 ? (
                              <span className="text-slate-600 italic">(Empty column)</span>
                            ) : (
                              samples.map((s, sIdx) => <div key={sIdx} className="truncate">• {s}</div>)
                            )}
                          </div>

                          <div className="space-y-1">
                            <label className="text-[9px] text-slate-500 font-bold uppercase block">Maps to System Field</label>
                            <select
                              value={currentMapping.systemField}
                              disabled={isSBHImport}
                              onChange={(e) => {
                                setColumnMappings(prev => ({
                                  ...prev,
                                  [colIdx]: {
                                    systemField: e.target.value,
                                    confidence: e.target.value === 'ignore' ? 0 : 100, // manual alignment gets high score
                                    matchedBy: 'manual'
                                  }
                                }));
                              }}
                              className={`w-full bg-slate-900 border border-slate-800 focus:border-indigo-500 rounded-lg px-3 py-1.5 text-xs text-slate-300 outline-none ${isSBHImport ? 'opacity-60 cursor-not-allowed' : ''}`}
                            >
                              <option value="ignore">🗑️ Skip Column / Ignore Data</option>
                              {mappingDictionary.map(entry => (
                                <option key={entry.field} value={entry.field}>
                                  {entry.isMandatory ? '⭐ ' : ''}{entry.label} ({entry.field})
                                </option>
                              ))}
                            </select>
                            <span className="text-[9px] text-slate-500 italic block font-mono">
                              {currentMapping.systemField !== 'ignore' && `Matched via: ${currentMapping.matchedBy}`}
                            </span>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}

              {/* STEP 3: PREVIEW DETAILED PARSED RECORDS WITH FORMATTING */}
              {importWizardStep === 'preview' && (
                <div className="space-y-5">
                  <div className="flex justify-between items-center flex-wrap gap-4">
                    <div>
                      <h4 className="font-bold text-slate-200 text-sm">Review Processed Rows ({excelRows.length} candidates)</h4>
                      <p className="text-xs text-slate-400 mt-0.5">
                        Double-check parsed fields. Row validation, automatic weight formatting, and defaults are applied below.
                      </p>
                    </div>
                    <div className="flex items-center gap-2">
                      <button
                        onClick={() => setImportWizardStep('mapping')}
                        className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg text-xs font-semibold"
                      >
                        Back
                      </button>
                      <button
                        onClick={executeIntelligentImport}
                        className="px-6 py-2 bg-emerald-600 hover:bg-emerald-500 text-white font-bold rounded-lg text-xs transition shadow-lg shadow-emerald-950/30 flex items-center gap-1.5 animate-pulse"
                      >
                        <CheckCircle className="w-4 h-4" />
                        Execute Database Import Now
                      </button>
                    </div>
                  </div>

                  {/* Field warnings legend if any */}
                  <div className="bg-slate-950 p-4 rounded-xl border border-slate-800 space-y-2">
                    <h5 className="text-xs font-bold text-slate-300 uppercase flex items-center gap-1">
                      <AlertCircle className="w-3.5 h-3.5 text-amber-500" />
                      Validation Warnings & Format Auto-Alignment
                    </h5>
                    <ul className="text-[11px] text-slate-400 space-y-1 list-disc list-inside leading-relaxed">
                      <li><strong>Name missing</strong>: Automatically replaced with <span className="font-mono bg-slate-900 px-1 text-slate-300">UNKNOWN PASSENGER</span> to prevent database schema errors.</li>
                      <li><strong>PIR missing</strong>: Replaced with <span className="font-mono bg-slate-900 px-1 text-slate-300">NO PIR</span> so records can load in reconciliation filters.</li>
                      <li><strong>Weight values</strong>: Stripped of units like &quot;kg&quot; or &quot;lbs&quot;, mapped as float decimals (left blank if not mentioned or invalid).</li>
                    </ul>
                  </div>

                  {/* Manifest Pre-view Table */}
                  <div className="bg-slate-950 rounded-xl border border-slate-800 overflow-hidden">
                    <div className="overflow-x-auto max-h-[400px]">
                      <table className="w-full text-left text-xs border-collapse">
                        <thead>
                          <tr className="bg-slate-900 text-slate-400 uppercase tracking-wider font-mono border-b border-slate-800">
                            <th className="py-2.5 px-4 font-semibold">Row</th>
                            <th className="py-2.5 px-3">Passenger Info (PIR / Name)</th>
                            <th className="py-2.5 px-3">Baggage Tags</th>
                            <th className="py-2.5 px-3">Flight & Dest</th>
                            <th className="py-2.5 px-3">Weight (kg)</th>
                            <th className="py-2.5 px-3 text-center">Damaged</th>
                            <th className="py-2.5 px-3 text-center">Locked</th>
                            <th className="py-2.5 px-3">Protocol</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-800/40 text-slate-300">
                          {excelRows.map((row, rIdx) => {
                            // Map values inline for quick preview UI
                            const getMappedVal = (field: string) => {
                              const colIdx = Object.entries(columnMappings).find(([_, mapping]) => mapping.systemField === field)?.[0];
                              if (colIdx !== undefined) {
                                  return row[Number(colIdx)] !== undefined ? String(row[Number(colIdx)]).trim() : '';
                              }
                              return '';
                            };

                            const pir = getMappedVal('pir');
                            const name = getMappedVal('name');
                            const originalTag = getMappedVal('originalTag');
                            const rushTag = getMappedVal('rushTag');
                            const flight = getMappedVal('flightNo') || 'LH760';
                            const dest = getMappedVal('destination');
                            const weightRaw = getMappedVal('weight');
                            const damaged = getMappedVal('damaged');
                            const rawLocked = getMappedVal('ln');
                            const protocol = getMappedVal('protocol');

                            // Compute weight display matching the executed import
                            let weightStr = '-';
                            if (weightRaw !== '') {
                              const cleanWeight = weightRaw.replace(/(kg|kgs|lbs|lb)\s*$/i, '').trim();
                              const parsed = Number(cleanWeight);
                              if (!isNaN(parsed) && parsed >= 0) {
                                weightStr = `${parsed} kg`;
                              }
                            }

                            // Compute lock dictionary expansion
                            let lockedDisp = rawLocked;
                            if (rawLocked !== '') {
                              const keys = Object.keys(lockDictionary);
                              const matchedKey = keys.find(k => k.toLowerCase() === rawLocked.toLowerCase());
                              if (matchedKey) {
                                lockedDisp = lockDictionary[matchedKey];
                              }
                            }

                            return (
                              <tr key={rIdx} className="hover:bg-slate-900/40">
                                <td className="py-2 px-4 text-slate-500 font-mono text-[10px]">{rIdx + 1}</td>
                                <td className="py-2 px-3">
                                  <div className="font-semibold text-slate-200 uppercase truncate max-w-[150px]">
                                    {name || <span className="text-amber-500 italic">UNKNOWN PASSENGER</span>}
                                  </div>
                                  <div className="text-[10px] font-mono text-slate-400">
                                    PIR: {pir || <span className="text-amber-500">NO PIR</span>}
                                  </div>
                                </td>
                                <td className="py-2 px-3 space-y-0.5">
                                  {originalTag && <div className="text-[11px] font-mono text-slate-300">Tag: {originalTag}</div>}
                                  {rushTag && (
                                    <div className={rushTag === 'Not Loaded' ? "text-[10px] font-mono text-red-500 font-bold" : "text-[10px] font-mono text-indigo-400 font-bold"}>
                                      {rushTag === 'Not Loaded' ? rushTag : `Rush: ${rushTag}`}
                                    </div>
                                  )}
                                  {!originalTag && !rushTag && <span className="text-slate-600">-</span>}
                                </td>
                                <td className="py-2 px-3">
                                  <span className="font-mono bg-slate-900 px-1.5 py-0.5 rounded text-indigo-300 font-bold">{flight}</span>
                                  <span className="text-slate-500 ml-1">→ {dest}</span>
                                </td>
                                <td className="py-2 px-3 font-mono text-indigo-300">
                                  {weightStr}
                                </td>
                                <td className="py-2 px-3 text-center">
                                  {/^(y|yes|true|dmg|damaged|1)$/i.test(damaged) ? (
                                    <span className="text-red-400 font-bold font-mono">Y</span>
                                  ) : (
                                    <span className="text-slate-500">N</span>
                                  )}
                                </td>
                                <td className="py-2 px-3 text-center">
                                  {lockedDisp ? (
                                    <span className={
                                      /^(n|no|false)$/i.test(lockedDisp)
                                        ? 'text-slate-500 font-bold'
                                        : 'text-red-400 font-bold font-mono'
                                    }>
                                      {lockedDisp}
                                    </span>
                                  ) : (
                                    <span className="text-slate-600">-</span>
                                  )}
                                </td>
                                <td className="py-2 px-3">
                                  <span className="text-[10px] bg-slate-900 px-2 py-0.5 rounded text-slate-400">
                                    {/non|other|uncleared|customs|cwc/i.test(protocol) ? 'Non-Cleared' : 'Cleared'}
                                  </span>
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>
              )}

              {/* STEP 4: SUMMARY RESULTS & DETAILED ERROR LOGS */}
              {importWizardStep === 'summary' && importSummaryResult && (
                <div className="space-y-6">
                  <div className="text-center p-6 bg-slate-950 rounded-2xl border border-slate-800 space-y-4 max-w-xl mx-auto">
                    <div className="w-12 h-12 bg-emerald-500/10 text-emerald-400 rounded-full flex items-center justify-center mx-auto">
                      <CheckCircle className="w-6 h-6" />
                    </div>
                    <div>
                      <h4 className="font-bold text-slate-200 text-lg">{isAlteaImport ? 'Altea Import' : 'Spreadsheet Import'} Completed Successfully</h4>
                      <p className="text-xs text-slate-400 mt-1">Swiss reconciliation database state synchronized</p>
                    </div>

                    <div className="grid grid-cols-3 gap-2 text-xs pt-2">
                      <div className="bg-slate-900 p-3 rounded-xl border border-slate-800">
                        <span className="text-2xl font-bold text-slate-100 block font-mono">{importSummaryResult.imported}</span>
                        <span className="text-[10px] text-slate-500 uppercase">Bags Imported</span>
                      </div>
                      <div className="bg-slate-900 p-3 rounded-xl border border-slate-800">
                        <span className="text-2xl font-bold text-slate-100 block font-mono">{importSummaryResult.duplicates}</span>
                        <span className="text-[10px] text-slate-500 uppercase font-medium">Matches Handled</span>
                      </div>
                      <div className="bg-slate-900 p-3 rounded-xl border border-slate-800">
                        <span className="text-2xl font-bold text-slate-100 block font-mono">{importSummaryResult.skipped}</span>
                        <span className="text-[10px] text-slate-500 uppercase">Rows Skipped</span>
                      </div>
                    </div>

                    {importSummaryResult.filesProcessed && (
                      <div className="bg-indigo-950/20 p-3 rounded-xl border border-indigo-500/20 text-indigo-200 text-[10px] font-mono flex justify-between items-center">
                        <span>Batch Files Processed:</span>
                        <span className="font-bold text-indigo-400">{importSummaryResult.filesProcessed} Files</span>
                      </div>
                    )}

                    {/* Secondary Data Quality Metrics */}
                    <div className="bg-slate-900/60 p-4 rounded-xl border border-slate-800/80 text-left space-y-3">
                      <h5 className="text-[10px] font-bold text-slate-400 uppercase tracking-wider border-b border-slate-800 pb-1.5">
                        Data Quality & Enrichment Metrics
                      </h5>
                      <div className="grid grid-cols-2 gap-4 text-xs font-mono">
                        <div className="space-y-1.5">
                          <div className="flex justify-between items-center text-slate-300">
                            <span>Blank Weights:</span>
                            <span className="text-slate-400 font-bold">{importSummaryResult.blankWeights}</span>
                          </div>
                          <div className="flex justify-between items-center text-slate-300">
                            <span>Invalid Weights:</span>
                            <span className={importSummaryResult.invalidWeights > 0 ? "text-amber-400 font-bold" : "text-slate-400 font-bold"}>
                              {importSummaryResult.invalidWeights}
                            </span>
                          </div>
                        </div>
                        <div className="space-y-1.5">
                          <div className="flex justify-between items-center text-slate-300">
                            <span>Recognized Locks:</span>
                            <span className="text-indigo-400 font-bold">{importSummaryResult.recognizedLocks}</span>
                          </div>
                          <div className="flex justify-between items-center text-slate-300">
                            <span>Unrecognized Locks:</span>
                            <span className="text-emerald-400 font-bold">{importSummaryResult.unrecognizedLocks}</span>
                          </div>
                        </div>
                      </div>
                    </div>

                    <button
                      onClick={() => {
                        setShowImportDialog(false);
                        setImportWizardStep('upload');
                        setExcelRows([]);
                        setExcelHeaders([]);
                        setRawPasteText('');
                        setImportSummaryResult(null);
                        setActiveSection('dashboard');
                        setIsSBHImport(false);
                        setIsAlteaImport(false);
                        setImportProgressMessage('');
                      }}
                      className="w-full bg-indigo-600 hover:bg-indigo-500 text-white font-bold py-2.5 rounded-xl text-xs transition"
                    >
                      Return to Dashboard
                    </button>
                  </div>

                  {/* Warnings and format alignments logs */}
                  {importSummaryResult.warnings.length > 0 && (
                    <div className="bg-slate-950 p-5 rounded-2xl border border-slate-800 space-y-3">
                      <div className="flex items-center gap-2 border-b border-slate-800 pb-2">
                        <AlertTriangle className="w-4 h-4 text-amber-500" />
                        <h5 className="text-xs font-bold text-slate-300 uppercase tracking-wider">Detailed Warnings & Format Logs</h5>
                      </div>
                      <div className="max-h-[150px] overflow-y-auto font-mono text-[10px] text-slate-400 space-y-1.5 divide-y divide-slate-800/30">
                        {importSummaryResult.warnings.map((warning, wIdx) => (
                          <div key={wIdx} className="pt-1 text-amber-300/80">• {warning}</div>
                        ))}
                      </div>
                    </div>
                  )}
                </div>
              )}

            </div>
          </div>
        </div>
      )}

      {/* SBH Import Instructions Modal */}
      {showSBHInstructions && (
        <div className="fixed inset-0 bg-black/90 backdrop-blur-md z-[100] flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-800 w-full max-w-2xl rounded-2xl shadow-2xl overflow-hidden flex flex-col max-h-[90vh] animate-in fade-in zoom-in duration-300">
            <div className="p-6 border-b border-slate-800 flex justify-between items-center bg-slate-900/50">
              <div className="flex items-center gap-3">
                <div className="p-2 bg-amber-500/20 rounded-lg">
                  <HelpCircle className="w-6 h-6 text-amber-400" />
                </div>
                <div>
                  <h3 className="font-bold text-slate-100 text-lg">Import Rush Bag List (SBH)</h3>
                  <p className="text-xs text-slate-400">Step-by-step instructions for Star Alliance Baggage Hub</p>
                </div>
              </div>
              <button onClick={() => setShowSBHInstructions(false)} className="p-2 hover:bg-slate-800 rounded-full transition">
                <X className="w-5 h-5 text-slate-500" />
              </button>
            </div>

            <div className="p-8 overflow-y-auto space-y-6">
              <div className="space-y-4">
                <div className="flex gap-4">
                  <div className="flex-shrink-0 w-8 h-8 bg-indigo-600 rounded-full flex items-center justify-center font-bold text-white shadow-lg shadow-indigo-600/20">1</div>
                  <div className="space-y-2">
                    <p className="text-slate-200 font-semibold">Open Star Alliance Baggage Hub (SBH)</p>
                    <a 
                      href="https://lhbaghub.staralliance.com/" 
                      target="_blank" 
                      rel="noopener noreferrer"
                      className="inline-flex items-center gap-2 bg-slate-950 border border-slate-800 hover:border-indigo-500/50 text-indigo-400 px-4 py-2 rounded-xl text-sm font-bold transition group"
                    >
                      lhbaghub.staralliance.com
                      <PlusCircle className="w-4 h-4 group-hover:rotate-45 transition-transform" />
                    </a>
                  </div>
                </div>

                <div className="flex gap-4 pt-2">
                  <div className="flex-shrink-0 w-8 h-8 bg-indigo-600 rounded-full flex items-center justify-center font-bold text-white shadow-lg shadow-indigo-600/20">2</div>
                  <div className="space-y-3">
                    <p className="text-slate-200 font-semibold">Generate the Rush Export Report</p>
                    <ul className="space-y-3">
                      <li className="flex gap-2 text-slate-400 text-sm">
                        <div className="w-1 h-1 bg-indigo-500 rounded-full mt-2 flex-shrink-0" />
                        <span>Navigate to: <strong className="text-slate-200">Search → Baggage Search</strong></span>
                      </li>
                      <li className="flex gap-2 text-slate-400 text-sm">
                        <div className="w-1 h-1 bg-indigo-500 rounded-full mt-2 flex-shrink-0" />
                        <span>Enter <strong className="text-slate-200">Date of Departure</strong> from HIN and <strong className="text-slate-200">Handling Station</strong> (FRA, MUC, or ZRH)</span>
                      </li>
                      <li className="flex gap-2 text-slate-400 text-sm">
                        <div className="w-1 h-1 bg-indigo-500 rounded-full mt-2 flex-shrink-0" />
                        <span>Enter <strong className="text-slate-200">Carrier & Flight Number</strong> (e.g. LH 760)</span>
                      </li>
                      <li className="flex gap-2 text-slate-400 text-sm">
                        <div className="w-1 h-1 bg-indigo-500 rounded-full mt-2 flex-shrink-0" />
                        <span>Select <strong className="text-slate-200">Bag Type</strong> as <strong className="text-slate-200">RUSH</strong> and <strong className="text-slate-200">WT Rush</strong></span>
                      </li>
                      <li className="flex gap-2 text-slate-400 text-sm italic bg-indigo-500/5 p-2 rounded border border-indigo-500/10">
                        <span>Click <strong className="text-slate-200">Search</strong>, then click <strong className="text-indigo-400">Export Excel</strong>. The file will be in your Downloads folder.</span>
                      </li>
                    </ul>
                  </div>
                </div>

                <div className="flex gap-4 pt-2">
                  <div className="flex-shrink-0 w-8 h-8 bg-indigo-600 rounded-full flex items-center justify-center font-bold text-white shadow-lg shadow-indigo-600/20">3</div>
                  <div className="space-y-2">
                    <p className="text-slate-200 font-semibold">Import the downloaded file</p>
                    <p className="text-slate-400 text-sm">Click the button below to select the Excel file from your computer.</p>
                  </div>
                </div>
              </div>

              <div className="pt-6 border-t border-slate-800 flex flex-col gap-3">
                <div className="flex flex-col sm:flex-row gap-3">
                  <button
                    onClick={() => {
                      setShowSBHInstructions(false);
                      setOcrTargetMode('table');
                      cameraInputRef.current?.click();
                    }}
                    className="flex-1 bg-amber-600 hover:bg-amber-500 text-white font-bold py-4 rounded-2xl flex items-center justify-center gap-3 transition shadow-lg shadow-amber-600/20 group"
                  >
                    <Camera className="w-5 h-5 group-hover:scale-110 transition-transform" />
                    Capture Sheet with Camera
                  </button>
                  <button
                    onClick={() => {
                      setShowSBHInstructions(false);
                      setIsSBHImport(true);
                      document.getElementById('sbh-file-input')?.click();
                    }}
                    className="flex-1 bg-indigo-600 hover:bg-indigo-500 text-white font-bold py-4 rounded-2xl flex items-center justify-center gap-3 transition shadow-lg shadow-indigo-600/20 group"
                  >
                    <Check className="w-5 h-5 group-hover:scale-125 transition-transform" />
                    Import Excel / List
                  </button>
                </div>
                <button
                  onClick={() => setShowSBHInstructions(false)}
                  className="w-full bg-slate-800 hover:bg-slate-700 text-slate-300 font-bold py-3 rounded-xl transition text-sm"
                >
                  Cancel
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Altea Import Instructions Modal */}
      {showAlteaInstructions && (
        <div className="fixed inset-0 bg-black/90 backdrop-blur-md z-[100] flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-800 w-full max-w-2xl rounded-2xl shadow-2xl overflow-hidden flex flex-col max-h-[90vh] animate-in fade-in zoom-in duration-300">
            <div className="p-6 border-b border-slate-800 flex justify-between items-center bg-slate-900/50">
              <div className="flex items-center gap-3">
                <div className="p-2 bg-emerald-500/20 rounded-lg">
                  <ClipboardList className="w-6 h-6 text-emerald-400" />
                </div>
                <div>
                  <h3 className="font-bold text-slate-100 text-lg">Import Rush Bags from Altea</h3>
                  <p className="text-xs text-slate-400">Step-by-step instructions for Altea exports</p>
                </div>
              </div>
              <button onClick={() => setShowAlteaInstructions(false)} className="p-2 hover:bg-slate-800 rounded-full transition">
                <X className="w-5 h-5 text-slate-500" />
              </button>
            </div>

            <div className="p-8 overflow-y-auto space-y-6">
              <div className="space-y-4">
                <p className="text-slate-300 text-sm">Go to: <strong className="text-slate-100">Flight → Baggage → Bag List → Rush Bags</strong></p>
                
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div className="bg-slate-950/50 p-4 rounded-xl border border-slate-800 space-y-2">
                    <h4 className="text-xs font-bold text-emerald-400 uppercase tracking-wider">Option 1 (Recommended)</h4>
                    <p className="text-[11px] font-bold text-slate-200">Device → Cryptic Output to Clipboard</p>
                    <p className="text-[10px] text-slate-400 leading-relaxed">
                      This option additionally provides the <strong className="text-slate-200">Loading Status</strong> (Loaded / Not Loaded), 
                      although baggage tags may occasionally contain minor OCR/copying inconsistencies.
                    </p>
                  </div>
                  <div className="bg-slate-950/50 p-4 rounded-xl border border-slate-800 space-y-2">
                    <h4 className="text-xs font-bold text-slate-400 uppercase tracking-wider">Option 2</h4>
                    <p className="text-[11px] font-bold text-slate-200">Device → Content to Clipboard</p>
                    <p className="text-[10px] text-slate-400 leading-relaxed">
                      This provides a cleaner list of baggage tags and passenger names but does not include loading status.
                    </p>
                  </div>
                </div>

                <div className="bg-indigo-950/20 p-4 rounded-xl border border-indigo-500/20 space-y-3">
                  <p className="text-xs text-indigo-200 font-semibold">After copying the contents:</p>
                  <ol className="text-[11px] text-indigo-300/80 space-y-2 list-decimal list-inside">
                    <li>Return to this application.</li>
                    <li>Click <strong>OK & Understood</strong>.</li>
                    <li>Paste the copied contents into the large text box.</li>
                    <li>Click <strong>Import</strong>.</li>
                  </ol>
                </div>
              </div>

              <div className="pt-6 border-t border-slate-800 flex flex-col gap-3">
                <div className="flex flex-col sm:flex-row gap-3">
                  <button
                    onClick={() => {
                      setShowAlteaInstructions(false);
                      setOcrTargetMode('altea');
                      cameraInputRef.current?.click();
                    }}
                    className="flex-1 bg-indigo-600 hover:bg-indigo-500 text-white font-bold py-4 rounded-2xl flex items-center justify-center gap-3 transition shadow-lg shadow-indigo-600/20 group"
                  >
                    <Camera className="w-5 h-5 group-hover:scale-110 transition-transform" />
                    Capture Altea Screen
                  </button>
                  <button
                    onClick={() => {
                      setShowAlteaInstructions(false);
                      setShowAlteaPasteWindow(true);
                    }}
                    className="flex-1 bg-emerald-600 hover:bg-emerald-500 text-white font-bold py-4 rounded-2xl flex items-center justify-center gap-3 transition shadow-lg shadow-emerald-600/20 group"
                  >
                    <Check className="w-5 h-5 group-hover:scale-125 transition-transform" />
                    Paste Clipboard
                  </button>
                </div>
                <button
                  onClick={() => setShowAlteaInstructions(false)}
                  className="w-full bg-slate-800 hover:bg-slate-700 text-slate-300 font-bold py-3 rounded-xl transition text-sm"
                >
                  Cancel
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Altea Paste Window Modal */}
      {showAlteaPasteWindow && (
        <div className="fixed inset-0 bg-black/95 backdrop-blur-md z-[110] flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-800 w-full max-w-5xl rounded-2xl shadow-2xl overflow-hidden flex flex-col h-[85vh] animate-in fade-in slide-in-from-bottom-8 duration-500">
            <div className="p-5 border-b border-slate-800 flex justify-between items-center bg-slate-900/50">
              <div className="flex items-center gap-3">
                <div className="p-2 bg-emerald-500/20 rounded-lg">
                  <ClipboardList className="w-5 h-5 text-emerald-400" />
                </div>
                <div>
                  <h3 className="font-bold text-slate-100">Paste Altea Rush Bag List</h3>
                  <p className="text-[10px] text-slate-500 font-mono">Format A: Cryptic Output | Format B: Content Output</p>
                </div>
              </div>
              <button 
                onClick={() => {
                  setShowAlteaPasteWindow(false);
                  setRawAlteaText('');
                }} 
                className="p-2 hover:bg-slate-800 rounded-full transition"
              >
                <X className="w-5 h-5 text-slate-500" />
              </button>
            </div>

            <div className="flex-1 p-6 bg-slate-950 overflow-hidden flex flex-col">
              <textarea
                autoFocus
                className="flex-1 w-full bg-transparent border-none focus:ring-0 text-slate-300 font-mono text-xs leading-relaxed resize-none p-4"
                placeholder="Paste Altea output here (Ctrl+V)..."
                value={rawAlteaText}
                onChange={(e) => setRawAlteaText(e.target.value)}
              />
            </div>

            <div className="p-5 border-t border-slate-800 bg-slate-900/50 flex flex-col md:flex-row justify-between items-center gap-4">
              <div className="flex items-center gap-4">
                <div className="hidden lg:block">
                  <h4 className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Duplicate Protocol</h4>
                </div>
                <div className="flex items-center gap-2">
                  <button 
                    onClick={() => setDuplicateMode('skip')}
                    className={`px-3 py-1.5 rounded-lg border font-semibold font-mono text-[10px] transition ${
                      duplicateMode === 'skip' 
                        ? 'bg-indigo-600 border-indigo-500 text-white' 
                        : 'bg-slate-900 border-slate-800 text-slate-500 hover:text-slate-300'
                    }`}
                  >
                    Skip Duplicates
                  </button>
                  <button 
                    onClick={() => setDuplicateMode('update')}
                    className={`px-3 py-1.5 rounded-lg border font-semibold font-mono text-[10px] transition ${
                      duplicateMode === 'update' 
                        ? 'bg-amber-600 border-amber-500 text-white' 
                        : 'bg-slate-900 border-slate-800 text-slate-500 hover:text-slate-300'
                    }`}
                  >
                    Overwrite & Update
                  </button>
                </div>
              </div>
              <div className="flex items-center gap-3">
                <div className="flex gap-3 mr-3 pr-3 border-r border-slate-800">
                  <button
                    onClick={() => setRawAlteaText('')}
                    className="px-6 py-2.5 bg-slate-800 hover:bg-slate-700 text-slate-400 text-xs font-bold rounded-xl transition"
                  >
                    Clear
                  </button>
                  <button
                    onClick={() => {
                      setShowAlteaPasteWindow(false);
                      setRawAlteaText('');
                    }}
                    className="px-6 py-2.5 border border-slate-800 hover:bg-slate-800 text-slate-500 text-xs font-bold rounded-xl transition"
                  >
                    Cancel
                  </button>
                </div>
                <button
                  disabled={!rawAlteaText.trim()}
                  onClick={() => handleAlteaImport()}
                  className="px-10 py-3 bg-emerald-600 hover:bg-emerald-500 disabled:opacity-50 disabled:cursor-not-allowed text-white font-bold rounded-xl flex items-center gap-2 transition shadow-lg shadow-emerald-600/20"
                >
                  <ArrowRightCircle className="w-5 h-5" />
                  Import
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Altea Preview Modal */}
      {showAlteaPreview && (
        <div className="fixed inset-0 bg-black/95 backdrop-blur-md z-[120] flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-800 w-full max-w-7xl rounded-2xl shadow-2xl overflow-hidden flex flex-col h-[90vh] animate-in fade-in slide-in-from-bottom-8 duration-500">
            <div className="p-5 border-b border-slate-800 flex justify-between items-center bg-slate-900/50">
              <div className="flex items-center gap-3">
                <div className="p-2 bg-emerald-500/20 rounded-lg">
                  <Eye className="w-5 h-5 text-emerald-400" />
                </div>
                <div>
                  <h3 className="font-bold text-slate-100">Altea Import Preview</h3>
                  <p className="text-[10px] text-slate-500 font-mono uppercase tracking-widest">Verify & Validate records before committing</p>
                </div>
              </div>
              <div className="flex items-center gap-4">
                <div className="bg-slate-950 px-4 py-1.5 rounded-full border border-slate-800 flex items-center gap-4">
                  <div className="flex items-center gap-2">
                    <div className="w-2 h-2 rounded-full bg-emerald-500"></div>
                    <span className="text-[10px] font-bold text-slate-300">Ready: {alteaPreviewRecords.filter(r => !baggageList.some(b => b.originalTag === r.originalTag) && r.originalTag && r.name && r.name !== 'UNKNOWN').length}</span>
                  </div>
                  <div className="flex items-center gap-2 border-l border-slate-800 pl-4">
                    <div className="w-2 h-2 rounded-full bg-amber-500"></div>
                    <span className="text-[10px] font-bold text-slate-300">Duplicate: {alteaPreviewRecords.filter(r => baggageList.some(b => b.originalTag === r.originalTag)).length}</span>
                  </div>
                  <div className="flex items-center gap-2 border-l border-slate-800 pl-4">
                    <div className="w-2 h-2 rounded-full bg-red-500"></div>
                    <span className="text-[10px] font-bold text-slate-300">Invalid: {alteaPreviewRecords.filter(r => !r.originalTag || r.originalTag.length < 6).length}</span>
                  </div>
                </div>
                <button 
                  onClick={() => {
                    setShowAlteaPreview(false);
                    setAlteaPreviewRecords([]);
                    setSelectedAlteaRowIds(new Set());
                  }} 
                  className="p-2 hover:bg-slate-800 rounded-full transition"
                >
                  <X className="w-5 h-5 text-slate-500" />
                </button>
              </div>
            </div>

            <div className="p-4 border-b border-slate-800 bg-slate-900/30 flex justify-between items-center gap-4">
              <div className="flex-1 max-w-md relative">
                <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-500" />
                <input
                  type="text"
                  placeholder="Search preview (Tag, Name, Flight...)"
                  value={alteaSearchQuery}
                  onChange={(e) => setAlteaSearchQuery(e.target.value)}
                  className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-2 pl-9 pr-4 rounded-xl focus:ring-1 focus:ring-indigo-500 transition"
                />
              </div>
              <div className="flex items-center gap-3">
                <button 
                  onClick={() => {
                    const allIds = new Set(alteaPreviewRecords.map(r => r.id));
                    setSelectedAlteaRowIds(allIds);
                  }}
                  className="px-3 py-1.5 text-[10px] font-bold text-indigo-400 hover:text-indigo-300 transition uppercase tracking-wider"
                >
                  Select All
                </button>
                <button 
                  onClick={() => setSelectedAlteaRowIds(new Set())}
                  className="px-3 py-1.5 text-[10px] font-bold text-slate-500 hover:text-slate-300 transition uppercase tracking-wider"
                >
                  Deselect All
                </button>
              </div>
            </div>

            <div className="flex-1 overflow-auto bg-slate-950">
              <table className="w-full border-collapse">
                <thead className="sticky top-0 bg-slate-900 z-10">
                  <tr className="border-b border-slate-800">
                    <th className="py-3 px-4 text-left w-12">
                      <input 
                        type="checkbox"
                        checked={selectedAlteaRowIds.size === alteaPreviewRecords.length && alteaPreviewRecords.length > 0}
                        onChange={(e) => {
                          if (e.target.checked) {
                            setSelectedAlteaRowIds(new Set(alteaPreviewRecords.map(r => r.id)));
                          } else {
                            setSelectedAlteaRowIds(new Set());
                          }
                        }}
                        className="w-4 h-4 rounded border-slate-700 bg-slate-800 text-indigo-600 focus:ring-indigo-500"
                      />
                    </th>
                    {[
                      { field: 'status', label: 'Status' },
                      { field: 'flightNo', label: 'Flight' },
                      { field: 'receivedAt', label: 'Date' },
                      { field: 'originalTag', label: 'Original Tag' },
                      { field: 'name', label: 'Passenger Name' },
                      { field: 'rushTag', label: 'Rush Status' },
                      { field: 'remarks', label: 'Remarks' }
                    ].map((col) => (
                      <th 
                        key={col.field}
                        onClick={() => {
                          if (alteaSortField === col.field) {
                            setAlteaSortDirection(alteaSortDirection === 'asc' ? 'desc' : 'asc');
                          } else {
                            setAlteaSortField(col.field as any);
                            setAlteaSortDirection('asc');
                          }
                        }}
                        className="py-3 px-4 text-left text-[10px] font-bold text-slate-500 uppercase tracking-wider cursor-pointer hover:bg-slate-800 transition-colors"
                      >
                        <div className="flex items-center gap-1">
                          {col.label}
                          {alteaSortField === col.field && (
                            alteaSortDirection === 'asc' ? <ArrowUp className="w-3 h-3" /> : <ArrowDown className="w-3 h-3" />
                          )}
                        </div>
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-900">
                  {alteaPreviewRecords
                    .filter(r => {
                      if (!alteaSearchQuery) return true;
                      const q = alteaSearchQuery.toLowerCase();
                      return (r.originalTag || '').toLowerCase().includes(q) || 
                             (r.name || '').toLowerCase().includes(q) || 
                             (r.flightNo || '').toLowerCase().includes(q);
                    })
                    .sort((a, b) => {
                      let valA: any = a[alteaSortField as keyof BaggageRecord] || '';
                      let valB: any = b[alteaSortField as keyof BaggageRecord] || '';

                      if (alteaSortField === 'status') {
                        const getStatusValue = (rec: BaggageRecord) => {
                          const isDup = baggageList.some(b => b.originalTag === rec.originalTag);
                          const isInv = !rec.originalTag || rec.originalTag.length < 6;
                          if (isInv) return 3;
                          if (isDup) return 2;
                          if (!rec.name || rec.name === 'UNKNOWN') return 1;
                          return 0;
                        };
                        valA = getStatusValue(a);
                        valB = getStatusValue(b);
                      }

                      if (typeof valA === 'string') {
                        return alteaSortDirection === 'asc' ? valA.localeCompare(valB) : valB.localeCompare(valA);
                      }
                      return alteaSortDirection === 'asc' ? (valA > valB ? 1 : -1) : (valB > valA ? 1 : -1);
                    })
                    .map((record) => {
                      const isDuplicate = baggageList.some(b => b.originalTag === record.originalTag);
                      const isInvalidTag = !record.originalTag || record.originalTag.length < 6;
                      const isMissingName = !record.name || record.name === 'UNKNOWN';
                      
                      let statusLabel = 'Ready';
                      let statusColor = 'text-emerald-400 bg-emerald-500/10 border-emerald-500/20';
                      let statusIcon = <CheckCircle className="w-3 h-3" />;
                      let statusReason = 'Record is valid and ready for import';
                      
                      if (isInvalidTag) {
                        statusLabel = 'Invalid Tag';
                        statusColor = 'text-red-400 bg-red-500/10 border-red-500/20';
                        statusIcon = <AlertCircle className="w-3 h-3" />;
                        statusReason = 'Baggage tag format is unrecognized or too short';
                      } else if (isDuplicate) {
                        statusLabel = 'Already Exists';
                        statusColor = 'text-amber-400 bg-amber-500/10 border-amber-500/20';
                        statusIcon = <Copy className="w-3 h-3" />;
                        statusReason = 'Tag number already exists in the reconciliation database';
                      } else if (isMissingName) {
                        statusLabel = 'Missing Name';
                        statusColor = 'text-orange-400 bg-orange-500/10 border-orange-500/20';
                        statusIcon = <AlertTriangle className="w-3 h-3" />;
                        statusReason = 'Passenger name could not be extracted from Altea data';
                      }

                      return (
                        <tr key={record.id} className={`hover:bg-slate-900/50 transition-colors ${selectedAlteaRowIds.has(record.id) ? 'bg-indigo-500/5' : ''}`}>
                          <td className="py-2.5 px-4">
                            <input 
                              type="checkbox"
                              checked={selectedAlteaRowIds.has(record.id)}
                              onChange={() => {
                                const newIds = new Set(selectedAlteaRowIds);
                                if (newIds.has(record.id)) newIds.delete(record.id);
                                else newIds.add(record.id);
                                setSelectedAlteaRowIds(newIds);
                              }}
                              className="w-4 h-4 rounded border-slate-700 bg-slate-800 text-indigo-600 focus:ring-indigo-500"
                            />
                          </td>
                          <td className="py-2.5 px-4">
                            <div 
                              title={statusReason}
                              className={`inline-flex items-center gap-1.5 px-2 py-0.5 rounded-full border ${statusColor} text-[9px] font-bold uppercase tracking-tight cursor-help`}
                            >
                              {statusIcon}
                              {statusLabel}
                            </div>
                          </td>
                          <td className="py-2.5 px-4 text-xs font-bold text-slate-300 font-mono">{record.flightNo}</td>
                          <td className="py-2.5 px-4 text-xs text-slate-400 font-mono">
                            {record.receivedAt ? new Date(record.receivedAt).toLocaleDateString('en-GB', { day: '2-digit', month: 'short' }).toUpperCase() : '-'}
                          </td>
                          <td className="py-2.5 px-4 text-xs font-bold text-slate-200 font-mono">{record.originalTag}</td>
                          <td className="py-2.5 px-4 text-xs text-slate-300 font-medium">{record.name}</td>
                          <td className="py-2.5 px-4">
                            {record.rushTag === 'Not Loaded' ? (
                              <span className="text-[10px] font-bold text-red-500 bg-red-500/10 px-1.5 py-0.5 rounded border border-red-500/20 uppercase">Not Loaded</span>
                            ) : (
                              <span className="text-slate-600 text-[10px]">-</span>
                            )}
                          </td>
                          <td className="py-2.5 px-4 text-[10px] text-slate-500 italic">{record.remarks}</td>
                        </tr>
                      );
                    })}
                </tbody>
              </table>
            </div>

            <div className="p-5 border-t border-slate-800 bg-slate-900/50 flex flex-col md:flex-row justify-between items-center gap-4">
              <div className="flex flex-col">
                <span className="text-[10px] font-bold text-slate-500 uppercase tracking-widest">Selection Summary</span>
                <span className="text-sm font-bold text-indigo-400">
                  {selectedAlteaRowIds.size} of {alteaPreviewRecords.length} records selected
                </span>
              </div>
              <div className="flex gap-4">
                <button
                  onClick={() => {
                    setShowAlteaPreview(false);
                    setShowAlteaPasteWindow(true);
                  }}
                  className="px-6 py-2.5 bg-slate-800 hover:bg-slate-700 text-slate-400 text-xs font-bold rounded-xl transition flex items-center gap-2"
                >
                  <ArrowLeft className="w-4 h-4" />
                  Back to Paste
                </button>
                <button
                  onClick={() => {
                    setShowAlteaPreview(false);
                    setAlteaPreviewRecords([]);
                    setRawAlteaText('');
                  }}
                  className="px-6 py-2.5 border border-slate-800 hover:bg-slate-800 text-slate-500 text-xs font-bold rounded-xl transition"
                >
                  Cancel
                </button>
                <button
                  disabled={selectedAlteaRowIds.size === 0}
                  onClick={() => handleCommitAlteaImport()}
                  className="px-10 py-3 bg-emerald-600 hover:bg-emerald-500 disabled:opacity-50 disabled:cursor-not-allowed text-white font-bold rounded-xl flex items-center gap-2 transition shadow-lg shadow-emerald-600/20"
                >
                  <CheckCircle className="w-5 h-5" />
                  Commit Import
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Hidden file input for SBH Import */}
      <input 
        id="sbh-file-input"
        type="file" 
        accept=".xlsx,.xls" 
        multiple
        style={{ display: 'none' }}
        onChange={(e) => {
          setIsSBHImport(true);
          handleCSVUpload(e);
        }}
      />

      {/* Hidden camera input for OCR */}
      <input
        ref={cameraInputRef}
        id="camera-ocr-input"
        type="file"
        accept="image/*"
        capture="environment"
        className="hidden"
        onChange={handleCameraCapture}
      />

      {/* OCR Loading Overlay */}
      {isOcrLoading && (
        <div className="fixed inset-0 bg-slate-950/80 backdrop-blur-md z-[200] flex flex-col items-center justify-center">
          <motion.div
            animate={{ rotate: 360 }}
            transition={{ repeat: Infinity, duration: 2, ease: "linear" }}
            className="w-16 h-16 border-4 border-indigo-500/20 border-t-indigo-500 rounded-full mb-6"
          />
          <h3 className="text-xl font-bold text-white mb-2">Analyzing Image...</h3>
          <p className="text-slate-400 font-mono text-xs animate-pulse">Gemini AI is extracting baggage data...</p>
          <div className="mt-8 w-64 h-1 bg-slate-800 rounded-full overflow-hidden">
            <motion.div
              initial={{ width: "0%" }}
              animate={{ width: "100%" }}
              transition={{ duration: 5, ease: "easeInOut" }}
              className="h-full bg-indigo-500"
            />
          </div>
        </div>
      )}
      {editingRecord && (
        <div className="fixed inset-0 bg-black/70 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-800 w-full max-w-xl rounded-xl shadow-2xl overflow-hidden flex flex-col">
            
            <div className="p-5 border-b border-slate-800 flex justify-between items-center">
              <div className="flex items-center gap-2">
                <Edit className="w-5 h-5 text-indigo-400" />
                <h3 className="font-bold text-slate-100 text-sm">Update Baggage & Customs Records</h3>
              </div>
              <button type="button" onClick={() => handleOpenEditDialog(null)}>
                <X className="w-4 h-4 text-slate-400 hover:text-white" />
              </button>
            </div>

            <form onSubmit={handleEditSaveSubmit} className="p-6 space-y-4 overflow-y-auto max-h-[80vh]">
              
              <div className="bg-slate-950 p-4 rounded-lg border border-slate-800 space-y-2">
                <p className="text-[10px] font-bold text-slate-400 uppercase tracking-wide">Primary Baggage Metadata</p>
                <div className="grid grid-cols-2 gap-3 text-xs text-slate-300">
                  <div>Passenger Name: <strong className="text-slate-100">{editingRecord.name}</strong></div>
                  <div>PIR ID: <strong className="text-slate-100 font-mono">{editingRecord.pir || 'NO PIR'}</strong></div>
                  <div>Original Tag: <strong className="text-slate-100 font-mono">{editingRecord.originalTag || '-'}</strong></div>
                  <div>Rush Tag: <strong className={editingRecord.rushTag === 'Not Loaded' ? 'text-red-400 font-mono' : 'text-indigo-400 font-mono'}>{editingRecord.rushTag || '-'}</strong></div>
                  <div>Flight Ops: <strong className="text-slate-100 font-mono">{editingRecord.flightNo}</strong></div>
                  <div>Destination: <strong className="text-slate-100 font-mono">{editingRecord.destination}</strong></div>
                  <div className="col-span-2 mt-2">
                    <label className="text-[10px] font-bold text-slate-400 uppercase block mb-1">Date of Baggage Receipt (Manual Override)</label>
                    <input
                      type="datetime-local"
                      value={editingRecord.receivedAt ? new Date(new Date(editingRecord.receivedAt).getTime() - (new Date().getTimezoneOffset() * 60000)).toISOString().slice(0, 16) : ''}
                      onChange={(e) => setEditingRecord({
                        ...editingRecord,
                        receivedAt: e.target.value ? new Date(e.target.value).toISOString() : undefined
                      })}
                      className="w-full bg-slate-950 border border-slate-800 focus:border-indigo-500 text-slate-200 px-3 py-1.5 rounded text-xs font-mono"
                    />
                  </div>
                </div>
              </div>

              {/* Enhanced Operational Fields in Edit Modal */}
              <div className="grid grid-cols-3 gap-3 bg-slate-950/40 p-4 rounded-lg border border-slate-800/80">
                {/* 1. Locked Free Text Input */}
                <div>
                  <label className="text-[10px] font-bold text-slate-400 uppercase block mb-1">Locked</label>
                  <input
                    type="text"
                    placeholder="PAD, CL, Y, N"
                    value={editingRecord.ln || ''}
                    onChange={(e) => setEditingRecord({
                      ...editingRecord,
                      ln: e.target.value
                    })}
                    className="w-full bg-slate-950 border border-slate-800 focus:border-indigo-500 text-slate-200 px-2 py-1 rounded text-xs font-mono"
                  />
                </div>

                {/* 2. Weight */}
                <div>
                  <label className="text-[10px] font-bold text-slate-400 uppercase block mb-1">Weight</label>
                  <div className="relative">
                    <input
                      type="number"
                      step="0.1"
                      min="0"
                      placeholder="e.g. 20.5"
                      value={editingRecord.weight === undefined ? '' : editingRecord.weight}
                      onChange={(e) => setEditingRecord({
                        ...editingRecord,
                        weight: e.target.value === '' ? undefined : Number(e.target.value)
                      })}
                      className="w-full bg-slate-950 border border-slate-800 focus:border-indigo-500 text-slate-200 pl-2 pr-6 py-1 rounded text-xs font-mono"
                    />
                    <span className="absolute right-1.5 top-1/2 -translate-y-1/2 text-[9px] text-slate-500 font-bold font-sans">kg</span>
                  </div>
                </div>

                {/* 3. Damaged toggle */}
                <div>
                  <label className="text-[10px] font-bold text-slate-400 uppercase block mb-1">Damaged *</label>
                  <div className="flex gap-1.5">
                    <button
                      type="button"
                      onClick={() => setEditingRecord({...editingRecord, damaged: 'Y'})}
                      className={`flex-1 py-1 text-[10px] rounded font-bold border transition ${
                        editingRecord.damaged === 'Y'
                          ? 'bg-amber-600 border-amber-500 text-white'
                          : 'bg-slate-950 border-slate-800 text-slate-400'
                      }`}
                    >
                      Y
                    </button>
                    <button
                      type="button"
                      onClick={() => setEditingRecord({...editingRecord, damaged: 'N'})}
                      className={`flex-1 py-1 text-[10px] rounded font-bold border transition ${
                        editingRecord.damaged === 'N' || !editingRecord.damaged
                          ? 'bg-indigo-600 border-indigo-500 text-white'
                          : 'bg-slate-950 border-slate-800 text-slate-400'
                      }`}
                    >
                      N
                    </button>
                  </div>
                </div>
              </div>

              {/* Dynamic Content based on Registry Type */}
              {editingRecord.registryType === 'Departure' ? (
                /* Departure Specific Edit Workflow */
                <div className="space-y-4">
                  <div className="bg-indigo-950/20 p-4 rounded-lg border border-indigo-500/20 space-y-4">
                    <div className="grid grid-cols-2 gap-4">
                      {/* Forwarding Status */}
                      <div className="space-y-1.5">
                        <label className="text-[10px] font-bold text-indigo-400 uppercase block">Forwarding Status *</label>
                        <select
                          required
                          value={editingRecord.departureForwardingStatus || 'Awaiting Forwarding'}
                          onChange={(e) => setEditingRecord({
                            ...editingRecord,
                            departureForwardingStatus: e.target.value as any,
                            disposition: (e.target.value === 'Forwarding Sent' ? 'Forwarded' : 'Awaiting Forwarding') as any
                          })}
                          className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-2 px-3 rounded cursor-pointer font-semibold"
                        >
                          <option value="Awaiting Forwarding">Awaiting Forwarding</option>
                          <option value="Forwarding Sent">Forwarding Sent</option>
                        </select>
                      </div>

                      {/* Storage Location */}
                      <div className="space-y-1.5">
                        <label className="text-[10px] font-bold text-indigo-400 uppercase block">Storage Location *</label>
                        <select
                          required
                          value={editingRecord.departureStorageLocation || 'LHG Office'}
                          onChange={(e) => setEditingRecord({
                            ...editingRecord,
                            departureStorageLocation: e.target.value as any,
                            dispositionLocation: e.target.value as any
                          })}
                          className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-2 px-3 rounded cursor-pointer font-semibold"
                        >
                          <option value="LHG Office">LHG Office</option>
                          <option value="BMA">BMA</option>
                          <option value="Level 4">Level 4</option>
                        </select>
                      </div>
                    </div>

                    {/* Conditional Forwarding Fields */}
                    <AnimatePresence>
                      {editingRecord.departureForwardingStatus === 'Forwarding Sent' && (
                        <motion.div
                          initial={{ opacity: 0, height: 0 }}
                          animate={{ opacity: 1, height: 'auto' }}
                          exit={{ opacity: 0, height: 0 }}
                          className="grid grid-cols-2 gap-3 bg-slate-950/60 p-3 rounded border border-slate-800"
                        >
                          <div>
                            <label className="text-[9px] font-bold text-slate-400 uppercase">Forwarding Flight *</label>
                            <input
                              type="text"
                              required
                              placeholder="LH763"
                              value={editingRecord.forwardingFlightNo || ''}
                              onChange={(e) => setEditingRecord({...editingRecord, forwardingFlightNo: e.target.value.toUpperCase()})}
                              className="w-full bg-slate-900 border border-slate-800 text-slate-200 px-2 py-1.5 rounded text-xs"
                            />
                          </div>
                          <div>
                            <label className="text-[9px] font-bold text-slate-400 uppercase">Forwarding Date *</label>
                            <input
                              type="date"
                              required
                              value={editingRecord.forwardingDate || ''}
                              onChange={(e) => setEditingRecord({...editingRecord, forwardingDate: e.target.value})}
                              className="w-full bg-slate-900 border border-slate-800 text-slate-200 px-2 py-1.5 rounded text-xs"
                            />
                          </div>
                          <div>
                            <label className="text-[9px] font-bold text-slate-400 uppercase">Destination *</label>
                            <input
                              type="text"
                              required
                              placeholder="FRA"
                              value={editingRecord.forwardingDestination || ''}
                              onChange={(e) => setEditingRecord({...editingRecord, forwardingDestination: e.target.value.toUpperCase()})}
                              className="w-full bg-slate-900 border border-slate-800 text-slate-200 px-2 py-1.5 rounded text-xs"
                            />
                          </div>
                          <div>
                            <label className="text-[9px] font-bold text-slate-400 uppercase">Forwarded By *</label>
                            <input
                              type="text"
                              required
                              placeholder="Staff ID"
                              value={editingRecord.forwardedBy || ''}
                              onChange={(e) => setEditingRecord({...editingRecord, forwardedBy: e.target.value.toUpperCase()})}
                              className="w-full bg-slate-900 border border-slate-800 text-slate-200 px-2 py-1.5 rounded text-xs"
                            />
                          </div>
                          <div className="col-span-2">
                            <label className="text-[9px] font-bold text-slate-400 uppercase">Forwarding Remarks</label>
                            <input
                              type="text"
                              placeholder="Optional remarks..."
                              value={editingRecord.forwardingRemarks || ''}
                              onChange={(e) => setEditingRecord({...editingRecord, forwardingRemarks: e.target.value})}
                              className="w-full bg-slate-900 border border-slate-800 text-slate-200 px-2 py-1.5 rounded text-xs"
                            />
                          </div>
                        </motion.div>
                      )}
                    </AnimatePresence>

                    {/* Conditional Storage Fields */}
                    <AnimatePresence>
                      {editingRecord.departureStorageLocation === 'BMA' && (
                        <motion.div
                          initial={{ opacity: 0, height: 0 }}
                          animate={{ opacity: 1, height: 'auto' }}
                          exit={{ opacity: 0, height: 0 }}
                          className="bg-slate-950/60 p-3 rounded border border-slate-800"
                        >
                          <label className="text-[9px] font-bold text-slate-400 uppercase">BMA Storage Location *</label>
                          <input
                            type="text"
                            required
                            placeholder="Cage/Rack location..."
                            value={editingRecord.bmaStorageLocation || ''}
                            onChange={(e) => setEditingRecord({...editingRecord, bmaStorageLocation: e.target.value})}
                            className="w-full bg-slate-900 border border-slate-800 text-slate-200 px-2 py-1.5 rounded text-xs"
                          />
                        </motion.div>
                      )}
                    </AnimatePresence>
                    {(editingRecord.departureStorageLocation === 'LHG Office' || editingRecord.departureStorageLocation === 'Level 4') && (
                      <div className="bg-slate-950/60 p-3 rounded border border-slate-800">
                        <label className="text-[9px] font-bold text-slate-400 uppercase">Storage Remarks</label>
                        <input
                          type="text"
                          placeholder="Optional storage details..."
                          value={editingRecord.storageRemarks || ''}
                          onChange={(e) => setEditingRecord({...editingRecord, storageRemarks: e.target.value})}
                          className="w-full bg-slate-900 border border-slate-800 text-slate-200 px-2 py-1.5 rounded text-xs"
                        />
                      </div>
                    )}
                  </div>

                  <div className="space-y-1.5">
                    <label className="text-[10px] font-bold text-slate-400 uppercase">Left Behind Reason</label>
                    <textarea
                      rows={3}
                      placeholder="Late acceptance, Security hold, etc."
                      value={editingRecord.leftBehindReason || ''}
                      onChange={(e) => setEditingRecord({...editingRecord, leftBehindReason: e.target.value})}
                      className="w-full bg-slate-950 border border-slate-800 focus:border-indigo-500 text-slate-200 px-3 py-1.5 rounded text-xs"
                    />
                  </div>
                </div>
              ) : (
                /* Redesigned Arrival Specific Edit Workflow */
                <>
                <div className="space-y-6">
                  {/* Step 1: Customs Status */}
                  <div className="space-y-1.5 bg-slate-950/40 p-4 rounded-lg border border-slate-800/80">
                    <label className="text-[10px] font-bold text-indigo-400 uppercase block tracking-wider">Step 1: Customs Status *</label>
                    <select
                      required
                      value={editingRecord.customsStatus === 'Cleared' ? 'Cleared' : (editingRecord.customsStatus === 'Not Cleared' || editingRecord.customsStatus === 'Marked Preventive' ? 'Not Cleared / Marked Preventive' : '')}
                      onChange={(e) => {
                        const val = e.target.value;
                        setEditingRecord({
                          ...editingRecord,
                          customsStatus: val === 'Cleared' ? 'Cleared' : 'Not Cleared',
                          // Reset subsequent fields
                          preventiveReason: '',
                          disposition: 'Pending',
                          arrivalBelt: undefined,
                          handoverAirline: '',
                          handoverDateTime: '',
                          warehouseReference: '',
                          storageRemarks: '',
                          reexportFlight: '',
                          reexportDestination: '',
                          reexportDate: '',
                          deliveryAgent: undefined,
                          deliveryDateTime: '',
                          domesticForwarding: undefined,
                          forwardingFlightNo: '',
                          forwardingDate: '',
                          forwardingDestination: '',
                        });
                      }}
                      className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-sm py-2 px-3 rounded-lg cursor-pointer font-semibold focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 transition-all"
                    >
                      <option value="">-- Select Status --</option>
                      <option value="Cleared">Cleared</option>
                      <option value="Not Cleared / Marked Preventive">Not Cleared / Marked Preventive</option>
                    </select>
                  </div>

                  <AnimatePresence mode="wait">
                    {/* Step 2 for Not Cleared */}
                    {(editingRecord.customsStatus === 'Not Cleared' || editingRecord.customsStatus === 'Marked Preventive') && (
                      <motion.div
                        initial={{ opacity: 0, y: 10 }}
                        animate={{ opacity: 1, y: 0 }}
                        exit={{ opacity: 0, y: -10 }}
                        className="space-y-4"
                      >
                        <div className="space-y-1.5 bg-slate-950/40 p-4 rounded-lg border border-slate-800/80">
                          <label className="text-[10px] font-bold text-indigo-400 uppercase block tracking-wider">Step 2: Preventive Reason *</label>
                          <textarea
                            required
                            rows={2}
                            placeholder="e.g. Customs hold, Security examination, etc."
                            value={editingRecord.preventiveReason || ''}
                            onChange={(e) => setEditingRecord({ ...editingRecord, preventiveReason: e.target.value })}
                            className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-2 px-3 rounded-lg focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 transition-all"
                          />
                        </div>

                        {editingRecord.preventiveReason && (
                          <motion.div
                            initial={{ opacity: 0, y: 10 }}
                            animate={{ opacity: 1, y: 0 }}
                            className="space-y-4"
                          >
                            <div className="space-y-1.5 bg-slate-950/40 p-4 rounded-lg border border-slate-800/80">
                              <label className="text-[10px] font-bold text-indigo-400 uppercase block tracking-wider">Step 3: Disposition *</label>
                              <select
                                required
                                value={editingRecord.disposition || ''}
                                onChange={(e) => {
                                  const disp = e.target.value as any;
                                  setEditingRecord({
                                    ...editingRecord,
                                    disposition: disp,
                                    // Reset sub-fields
                                    arrivalBelt: undefined,
                                    handoverAirline: '',
                                    handoverDateTime: disp === 'Handover' ? new Date().toISOString().slice(0, 16) : '',
                                    warehouseReference: '',
                                    storageRemarks: '',
                                    reexportFlight: '',
                                    reexportDestination: '',
                                    reexportDate: '',
                                  });
                                }}
                                className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-sm py-2 px-3 rounded-lg cursor-pointer font-semibold focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 transition-all"
                              >
                                <option value="">-- Select Disposition --</option>
                                <option value="Belt 9">Arrival Belt 9</option>
                                <option value="Handover">Handed Over to OAL (Other Airline)</option>
                                <option value="CWC">CWC Warehouse</option>
                                <option value="Re-export">Re-export</option>
                              </select>
                            </div>

                            {/* Dynamic Disposition Fields for Not Cleared */}
                            <AnimatePresence mode="wait">
                              {editingRecord.disposition === 'Belt 9' && (
                                <motion.div
                                  key="belt9"
                                  initial={{ opacity: 0, height: 0 }}
                                  animate={{ opacity: 1, height: 'auto' }}
                                  exit={{ opacity: 0, height: 0 }}
                                  className="space-y-1.5 bg-slate-950/40 p-4 rounded-lg border border-slate-800/80"
                                >
                                  <label className="text-[10px] font-bold text-indigo-400 uppercase block">Remarks (Optional)</label>
                                  <input
                                    type="text"
                                    placeholder="Lounge holding area..."
                                    value={editingRecord.remarks || ''}
                                    onChange={(e) => setEditingRecord({ ...editingRecord, remarks: e.target.value })}
                                    className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-2 px-3 rounded-lg"
                                  />
                                </motion.div>
                              )}

                              {editingRecord.disposition === 'Handover' && (
                                <motion.div
                                  key="handover"
                                  initial={{ opacity: 0, height: 0 }}
                                  animate={{ opacity: 1, height: 'auto' }}
                                  exit={{ opacity: 0, height: 0 }}
                                  className="space-y-4 bg-slate-950/40 p-4 rounded-lg border border-slate-800/80"
                                >
                                  <div>
                                    <label className="text-[10px] font-bold text-indigo-400 uppercase block mb-1">Partner Airline *</label>
                                    <input
                                      type="text"
                                      required
                                      placeholder="Air India, Emirates, etc."
                                      value={editingRecord.handoverAirline || ''}
                                      onChange={(e) => setEditingRecord({ ...editingRecord, handoverAirline: e.target.value })}
                                      className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-2 px-3 rounded-lg"
                                    />
                                  </div>
                                  <div>
                                    <label className="text-[10px] font-bold text-indigo-400 uppercase block mb-1">Handover Date & Time *</label>
                                    <input
                                      type="datetime-local"
                                      required
                                      value={editingRecord.handoverDateTime || ''}
                                      onChange={(e) => setEditingRecord({ ...editingRecord, handoverDateTime: e.target.value })}
                                      className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-2 px-3 rounded-lg"
                                    />
                                  </div>
                                </motion.div>
                              )}

                              {editingRecord.disposition === 'CWC' && (
                                <motion.div
                                  key="cwc"
                                  initial={{ opacity: 0, height: 0 }}
                                  animate={{ opacity: 1, height: 'auto' }}
                                  exit={{ opacity: 0, height: 0 }}
                                  className="space-y-4 bg-slate-950/40 p-4 rounded-lg border border-slate-800/80"
                                >
                                  <div>
                                    <label className="text-[10px] font-bold text-indigo-400 uppercase block mb-1">Warehouse Reference *</label>
                                    <input
                                      type="text"
                                      required
                                      placeholder="Ref #12345"
                                      value={editingRecord.warehouseReference || ''}
                                      onChange={(e) => setEditingRecord({ ...editingRecord, warehouseReference: e.target.value })}
                                      className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-2 px-3 rounded-lg"
                                    />
                                  </div>
                                  <div>
                                    <label className="text-[10px] font-bold text-indigo-400 uppercase block mb-1">Storage Remarks (Optional)</label>
                                    <input
                                      type="text"
                                      placeholder="Central depot storage..."
                                      value={editingRecord.storageRemarks || ''}
                                      onChange={(e) => setEditingRecord({ ...editingRecord, storageRemarks: e.target.value })}
                                      className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-2 px-3 rounded-lg"
                                    />
                                  </div>
                                </motion.div>
                              )}

                              {editingRecord.disposition === 'Re-export' && (
                                <motion.div
                                  key="reexport"
                                  initial={{ opacity: 0, height: 0 }}
                                  animate={{ opacity: 1, height: 'auto' }}
                                  exit={{ opacity: 0, height: 0 }}
                                  className="grid grid-cols-2 gap-4 bg-slate-950/40 p-4 rounded-lg border border-slate-800/80"
                                >
                                  <div>
                                    <label className="text-[10px] font-bold text-indigo-400 uppercase block mb-1">Re-export Flight *</label>
                                    <input
                                      type="text"
                                      required
                                      placeholder="LH761"
                                      value={editingRecord.reexportFlight || ''}
                                      onChange={(e) => setEditingRecord({ ...editingRecord, reexportFlight: e.target.value.toUpperCase() })}
                                      className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-2 px-3 rounded-lg"
                                    />
                                  </div>
                                  <div>
                                    <label className="text-[10px] font-bold text-indigo-400 uppercase block mb-1">Destination *</label>
                                    <input
                                      type="text"
                                      required
                                      placeholder="FRA"
                                      value={editingRecord.reexportDestination || ''}
                                      onChange={(e) => setEditingRecord({ ...editingRecord, reexportDestination: e.target.value.toUpperCase() })}
                                      className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-2 px-3 rounded-lg"
                                    />
                                  </div>
                                  <div className="col-span-2">
                                    <label className="text-[10px] font-bold text-indigo-400 uppercase block mb-1">Re-export Date *</label>
                                    <input
                                      type="date"
                                      required
                                      value={editingRecord.reexportDate || ''}
                                      onChange={(e) => setEditingRecord({ ...editingRecord, reexportDate: e.target.value })}
                                      className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-2 px-3 rounded-lg"
                                    />
                                  </div>
                                  <div className="col-span-2">
                                    <label className="text-[10px] font-bold text-indigo-400 uppercase block mb-1">Remarks</label>
                                    <input
                                      type="text"
                                      placeholder="Repatriation details..."
                                      value={editingRecord.remarks || ''}
                                      onChange={(e) => setEditingRecord({ ...editingRecord, remarks: e.target.value })}
                                      className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-2 px-3 rounded-lg"
                                    />
                                  </div>
                                </motion.div>
                              )}
                            </AnimatePresence>
                          </motion.div>
                        )}
                      </motion.div>
                    )}

                    {/* Step 2 for Cleared */}
                    {editingRecord.customsStatus === 'Cleared' && (
                      <motion.div
                        initial={{ opacity: 0, y: 10 }}
                        animate={{ opacity: 1, y: 0 }}
                        exit={{ opacity: 0, y: -10 }}
                        className="space-y-4"
                      >
                        <div className="space-y-1.5 bg-slate-950/40 p-4 rounded-lg border border-slate-800/80">
                          <label className="text-[10px] font-bold text-indigo-400 uppercase block tracking-wider">Step 2: Disposition *</label>
                          <select
                            required
                            value={editingRecord.disposition || ''}
                            onChange={(e) => {
                              const disp = e.target.value as any;
                              setEditingRecord({
                                ...editingRecord,
                                disposition: disp,
                                // Reset sub-fields
                                deliveryAgent: disp === 'Delivered' ? 'VVM' : undefined,
                                deliveryDateTime: disp === 'Delivered' ? new Date().toISOString().slice(0, 16) : '',
                                domesticForwarding: disp === 'Forwarded' ? 'Air India' : undefined,
                                storageRemarks: '',
                                forwardingFlightNo: '',
                                forwardingDate: '',
                                forwardingDestination: '',
                              });
                            }}
                            className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-sm py-2 px-3 rounded-lg cursor-pointer font-semibold focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 transition-all"
                          >
                            <option value="">-- Select Disposition --</option>
                            <option value="Delivered">Delivery Agent</option>
                            <option value="Storage">LHG Office</option>
                            <option value="Forwarded">Domestic Forwarding</option>
                          </select>
                        </div>

                        {/* Dynamic Disposition Fields for Cleared */}
                        <AnimatePresence mode="wait">
                          {editingRecord.disposition === 'Delivered' && (
                            <motion.div
                              key="delivered"
                              initial={{ opacity: 0, height: 0 }}
                              animate={{ opacity: 1, height: 'auto' }}
                              exit={{ opacity: 0, height: 0 }}
                              className="space-y-4 bg-slate-950/40 p-4 rounded-lg border border-slate-800/80"
                            >
                              <div>
                                <label className="text-[10px] font-bold text-indigo-400 uppercase block mb-1">Delivery Agent *</label>
                                <select
                                  required
                                  value={editingRecord.deliveryAgent || 'VVM'}
                                  onChange={(e) => setEditingRecord({ ...editingRecord, deliveryAgent: e.target.value as any })}
                                  className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-2 px-3 rounded-lg"
                                >
                                  <option value="VVM">VVM</option>
                                  <option value="Outlook">Outlook</option>
                                  <option value="Advik">Advik</option>
                                </select>
                              </div>
                              <div>
                                <label className="text-[10px] font-bold text-indigo-400 uppercase block mb-1">Delivery Date & Time *</label>
                                <input
                                  type="datetime-local"
                                  required
                                  value={editingRecord.deliveryDateTime || ''}
                                  onChange={(e) => setEditingRecord({ ...editingRecord, deliveryDateTime: e.target.value })}
                                  className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-2 px-3 rounded-lg"
                                />
                              </div>
                              <div>
                                <label className="text-[10px] font-bold text-indigo-400 uppercase block mb-1">Remarks (Optional)</label>
                                <input
                                  type="text"
                                  placeholder="Direct dispatch details..."
                                  value={editingRecord.remarks || ''}
                                  onChange={(e) => setEditingRecord({ ...editingRecord, remarks: e.target.value })}
                                  className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-2 px-3 rounded-lg"
                                />
                              </div>
                            </motion.div>
                          )}

                          {editingRecord.disposition === 'Storage' && (
                            <motion.div
                              key="storage"
                              initial={{ opacity: 0, height: 0 }}
                              animate={{ opacity: 1, height: 'auto' }}
                              exit={{ opacity: 0, height: 0 }}
                              className="space-y-1.5 bg-slate-950/40 p-4 rounded-lg border border-slate-800/80"
                            >
                              <label className="text-[10px] font-bold text-indigo-400 uppercase block">Storage Remarks (Optional)</label>
                              <input
                                type="text"
                                placeholder="LHG Office storage details..."
                                value={editingRecord.storageRemarks || ''}
                                onChange={(e) => setEditingRecord({ ...editingRecord, storageRemarks: e.target.value })}
                                className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-2 px-3 rounded-lg"
                              />
                            </motion.div>
                          )}

                          {editingRecord.disposition === 'Forwarded' && (
                            <motion.div
                              key="forwarded"
                              initial={{ opacity: 0, height: 0 }}
                              animate={{ opacity: 1, height: 'auto' }}
                              exit={{ opacity: 0, height: 0 }}
                              className="grid grid-cols-2 gap-4 bg-slate-950/40 p-4 rounded-lg border border-slate-800/80"
                            >
                              <div className="col-span-2">
                                <label className="text-[10px] font-bold text-indigo-400 uppercase block mb-1">Forward Via *</label>
                                <select
                                  required
                                  value={editingRecord.domesticForwarding || 'Air India'}
                                  onChange={(e) => setEditingRecord({ ...editingRecord, domesticForwarding: e.target.value as any })}
                                  className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-2 px-3 rounded-lg"
                                >
                                  <option value="Air India">Air India</option>
                                  <option value="SpiceJet">SpiceJet</option>
                                  <option value="IndiGo">IndiGo</option>
                                </select>
                              </div>
                              <div>
                                <label className="text-[10px] font-bold text-indigo-400 uppercase block mb-1">Forwarding Flight Number *</label>
                                <input
                                  type="text"
                                  required
                                  placeholder="AI808"
                                  value={editingRecord.forwardingFlightNo || ''}
                                  onChange={(e) => setEditingRecord({ ...editingRecord, forwardingFlightNo: e.target.value.toUpperCase() })}
                                  className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-2 px-3 rounded-lg"
                                />
                              </div>
                              <div>
                                <label className="text-[10px] font-bold text-indigo-400 uppercase block mb-1">Forwarding Date *</label>
                                <input
                                  type="date"
                                  required
                                  value={editingRecord.forwardingDate || ''}
                                  onChange={(e) => setEditingRecord({ ...editingRecord, forwardingDate: e.target.value })}
                                  className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-2 px-3 rounded-lg"
                                />
                              </div>
                              <div className="col-span-2">
                                <label className="text-[10px] font-bold text-indigo-400 uppercase block mb-1">Destination *</label>
                                <input
                                  type="text"
                                  required
                                  placeholder="DEL"
                                  value={editingRecord.forwardingDestination || ''}
                                  onChange={(e) => setEditingRecord({ ...editingRecord, forwardingDestination: e.target.value.toUpperCase() })}
                                  className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-2 px-3 rounded-lg"
                                />
                              </div>
                              <div className="col-span-2">
                                <label className="text-[10px] font-bold text-indigo-400 uppercase block mb-1">Remarks</label>
                                <input
                                  type="text"
                                  placeholder="Domestic connection details..."
                                  value={editingRecord.remarks || ''}
                                  onChange={(e) => setEditingRecord({ ...editingRecord, remarks: e.target.value })}
                                  className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-2 px-3 rounded-lg"
                                />
                              </div>
                            </motion.div>
                          )}
                        </AnimatePresence>
                      </motion.div>
                    )}
                  </AnimatePresence>
                </div>

                  {/* Storage Location Remarks (requested feature) */}
                  <div className="space-y-1.5">
                    <label className="text-[10px] font-bold text-slate-400 uppercase">Storage Location Remarks</label>
                    <input
                      type="text"
                      placeholder="Specify rack, bin, shelf number, or other storage remarks..."
                      value={editingRecord.storageRemarks || ''}
                      onChange={(e) => setEditingRecord({
                        ...editingRecord,
                        storageRemarks: e.target.value
                      })}
                      className="w-full bg-slate-950 border border-slate-800 focus:border-indigo-500 text-slate-200 px-3 py-2 rounded text-xs"
                    />
                  </div>

                  {/* Manual backdate for test storage alert warnings */}
                  <div className="bg-slate-950 p-3 rounded border border-slate-800 space-y-2">
                    <label className="text-[10px] font-bold text-slate-400 uppercase block">Reconciliation Storage Clock (Testing Alerts)</label>
                    <div className="flex items-center gap-3">
                      <div className="flex-1">
                        <p className="text-[10px] text-slate-400">Received date / Storage Entry:</p>
                        <input
                          type="datetime-local"
                          value={editingRecord.receivedAt ? new Date(new Date(editingRecord.receivedAt).getTime() - (new Date().getTimezoneOffset() * 60000)).toISOString().slice(0, 16) : ''}
                          onChange={(e) => setEditingRecord({
                            ...editingRecord,
                            receivedAt: e.target.value ? new Date(e.target.value).toISOString() : undefined
                          })}
                          className="w-full bg-slate-900 border border-slate-800 text-slate-200 text-xs px-2 py-1.5 rounded"
                        />
                      </div>
                      <div className="flex flex-col gap-1.5">
                        <button
                          type="button"
                          onClick={() => {
                            const date3DaysAgo = new Date(Date.now() - 3.1 * 24 * 60 * 60 * 1000).toISOString();
                            setEditingRecord({
                              ...editingRecord,
                              receivedAt: date3DaysAgo,
                              disposition: 'Storage',
                              dispositionLocation: 'LHG Office'
                            });
                          }}
                          className="px-2 py-1 bg-amber-500/10 hover:bg-amber-500/20 text-amber-300 border border-amber-500/30 rounded text-[10px]"
                        >
                          Trigger 3d Storage warning
                        </button>
                        <button
                          type="button"
                          onClick={() => {
                            const date5DaysAgo = new Date(Date.now() - 5.1 * 24 * 60 * 60 * 1000).toISOString();
                            setEditingRecord({
                              ...editingRecord,
                              receivedAt: date5DaysAgo,
                              disposition: 'Storage',
                              dispositionLocation: 'CWC'
                            });
                          }}
                          className="px-2 py-1 bg-red-500/10 hover:bg-red-500/20 text-red-300 border border-red-500/30 rounded text-[10px]"
                        >
                          Trigger 5d Re-export urgency
                        </button>
                      </div>
                    </div>
                  </div>

                  {/* Arrival Reconciliation Remarks field (Moved outside status check for consistency) */}
                  <div className="space-y-1.5">
                    <label className="text-[10px] font-bold text-slate-400 uppercase">General Remarks</label>
                    <textarea
                      rows={2}
                      value={editingRecord.remarks || ''}
                      onChange={(e) => setEditingRecord({...editingRecord, remarks: e.target.value})}
                      className="w-full bg-slate-950 border border-slate-800 focus:border-indigo-500 text-slate-200 px-3 py-1.5 rounded text-xs"
                    />
                  </div>
                </>
              )}

              <div className="flex justify-end gap-2 pt-4 border-t border-slate-800">
                <button
                  type="button"
                  onClick={() => handleOpenEditDialog(null)}
                  className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded text-xs transition"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  className="px-5 py-2 bg-indigo-600 hover:bg-indigo-500 text-white rounded text-xs font-semibold transition"
                >
                  Save Changes
                </button>
              </div>
            </form>

          </div>
        </div>
      )}

      {/* C. BULK EDIT RECONCILIATION DIALOG */}
      {showBulkEdit && (
        <div className="fixed inset-0 bg-black/70 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-800 w-full max-w-lg rounded-xl shadow-2xl overflow-hidden flex flex-col">
            
            <div className="p-5 border-b border-slate-800 flex justify-between items-center bg-slate-900">
              <h3 className="font-bold text-slate-100 text-sm">Bulk Status & Disposition Processing</h3>
              <button onClick={() => setShowBulkEdit(false)}>
                <X className="w-4 h-4 text-slate-400 hover:text-white" />
              </button>
            </div>

            <form onSubmit={handleBulkEditSubmit} className="p-6 space-y-4 overflow-y-auto" style={{ maxHeight: '75vh' }}>
              <p className="text-[10px] text-indigo-300 bg-indigo-950/20 p-3 border border-indigo-900/30 rounded uppercase tracking-widest font-bold">
                Bulk Update: {selectedIds.length} Records Selected
              </p>

              {/* Step 1: Customs Status */}
              <div className="space-y-1.5 bg-slate-950/40 p-4 rounded-lg border border-slate-800/80">
                <label className="text-[10px] font-bold text-indigo-400 uppercase block tracking-wider">Step 1: Set Customs Status *</label>
                <select
                  required
                  value={bulkCustomsStatus === 'Cleared' ? 'Cleared' : (bulkCustomsStatus === 'Not Cleared' || bulkCustomsStatus === 'Marked Preventive' ? 'Not Cleared / Marked Preventive' : '')}
                  onChange={(e) => {
                    const val = e.target.value;
                    setBulkCustomsStatus(val === 'Cleared' ? 'Cleared' : 'Not Cleared');
                    // Reset subsequent fields
                    setBulkPreventiveReason('');
                    setBulkDisposition('Pending');
                  }}
                  className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-sm py-2 px-3 rounded-lg cursor-pointer font-semibold focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 transition-all"
                >
                  <option value="">-- Select Status --</option>
                  <option value="Cleared">Cleared</option>
                  <option value="Not Cleared / Marked Preventive">Not Cleared / Marked Preventive</option>
                </select>
              </div>

              <AnimatePresence mode="wait">
                {/* Step 2 for Not Cleared */}
                {(bulkCustomsStatus === 'Not Cleared' || bulkCustomsStatus === 'Marked Preventive') && (
                  <motion.div
                    initial={{ opacity: 0, y: 10 }}
                    animate={{ opacity: 1, y: 0 }}
                    exit={{ opacity: 0, y: -10 }}
                    className="space-y-4"
                  >
                    <div className="space-y-1.5 bg-slate-950/40 p-4 rounded-lg border border-slate-800/80">
                      <label className="text-[10px] font-bold text-indigo-400 uppercase block tracking-wider">Step 2: Preventive Reason *</label>
                      <textarea
                        required
                        rows={2}
                        placeholder="e.g. Customs hold, Security examination, etc."
                        value={bulkPreventiveReason}
                        onChange={(e) => setBulkPreventiveReason(e.target.value)}
                        className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-2 px-3 rounded-lg focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 transition-all"
                      />
                    </div>

                    {bulkPreventiveReason && (
                      <motion.div
                        initial={{ opacity: 0, y: 10 }}
                        animate={{ opacity: 1, y: 0 }}
                        className="space-y-4"
                      >
                        <div className="space-y-1.5 bg-slate-950/40 p-4 rounded-lg border border-slate-800/80">
                          <label className="text-[10px] font-bold text-indigo-400 uppercase block tracking-wider">Step 3: Disposition *</label>
                          <select
                            required
                            value={bulkDisposition}
                            onChange={(e) => {
                              const disp = e.target.value as any;
                              setBulkDisposition(disp);
                              if (disp === 'Handover') setBulkHandoverDateTime(new Date().toISOString().slice(0, 16));
                            }}
                            className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-sm py-2 px-3 rounded-lg cursor-pointer font-semibold focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 transition-all"
                          >
                            <option value="Pending">-- Select Disposition --</option>
                            <option value="Belt 9">Arrival Belt 9</option>
                            <option value="Handover">Handed Over to OAL (Other Airline)</option>
                            <option value="CWC">CWC Warehouse</option>
                            <option value="Re-export">Re-export</option>
                          </select>
                        </div>

                        {/* Dynamic Disposition Fields for Not Cleared */}
                        <AnimatePresence mode="wait">
                          {bulkDisposition === 'Belt 9' && (
                            <motion.div
                              key="bulkBelt9"
                              initial={{ opacity: 0, height: 0 }}
                              animate={{ opacity: 1, height: 'auto' }}
                              exit={{ opacity: 0, height: 0 }}
                              className="space-y-1.5 bg-slate-950/40 p-4 rounded-lg border border-slate-800/80"
                            >
                              <label className="text-[10px] font-bold text-indigo-400 uppercase block">Remarks (Optional)</label>
                              <input
                                type="text"
                                placeholder="Lounge holding area..."
                                value={bulkRemarks}
                                onChange={(e) => setBulkRemarks(e.target.value)}
                                className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-2 px-3 rounded-lg"
                              />
                            </motion.div>
                          )}

                          {bulkDisposition === 'Handover' && (
                            <motion.div
                              key="bulkHandover"
                              initial={{ opacity: 0, height: 0 }}
                              animate={{ opacity: 1, height: 'auto' }}
                              exit={{ opacity: 0, height: 0 }}
                              className="space-y-4 bg-slate-950/40 p-4 rounded-lg border border-slate-800/80"
                            >
                              <div>
                                <label className="text-[10px] font-bold text-indigo-400 uppercase block mb-1">Partner Airline *</label>
                                <input
                                  type="text"
                                  required
                                  placeholder="Air India, Emirates, etc."
                                  value={bulkHandoverAirline}
                                  onChange={(e) => setBulkHandoverAirline(e.target.value)}
                                  className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-2 px-3 rounded-lg"
                                />
                              </div>
                              <div>
                                <label className="text-[10px] font-bold text-indigo-400 uppercase block mb-1">Handover Date & Time *</label>
                                <input
                                  type="datetime-local"
                                  required
                                  value={bulkHandoverDateTime}
                                  onChange={(e) => setBulkHandoverDateTime(e.target.value)}
                                  className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-2 px-3 rounded-lg"
                                />
                              </div>
                            </motion.div>
                          )}

                          {bulkDisposition === 'CWC' && (
                            <motion.div
                              key="bulkCwc"
                              initial={{ opacity: 0, height: 0 }}
                              animate={{ opacity: 1, height: 'auto' }}
                              exit={{ opacity: 0, height: 0 }}
                              className="space-y-4 bg-slate-950/40 p-4 rounded-lg border border-slate-800/80"
                            >
                              <div>
                                <label className="text-[10px] font-bold text-indigo-400 uppercase block mb-1">Warehouse Reference *</label>
                                <input
                                  type="text"
                                  required
                                  placeholder="Ref #12345"
                                  value={bulkWarehouseReference}
                                  onChange={(e) => setBulkWarehouseReference(e.target.value)}
                                  className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-2 px-3 rounded-lg"
                                />
                              </div>
                              <div>
                                <label className="text-[10px] font-bold text-indigo-400 uppercase block mb-1">Storage Remarks (Optional)</label>
                                <input
                                  type="text"
                                  placeholder="Central depot storage..."
                                  value={bulkStorageRemarks}
                                  onChange={(e) => setBulkStorageRemarks(e.target.value)}
                                  className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-2 px-3 rounded-lg"
                                />
                              </div>
                            </motion.div>
                          )}

                          {bulkDisposition === 'Re-export' && (
                            <motion.div
                              key="bulkReexport"
                              initial={{ opacity: 0, height: 0 }}
                              animate={{ opacity: 1, height: 'auto' }}
                              exit={{ opacity: 0, height: 0 }}
                              className="grid grid-cols-2 gap-4 bg-slate-950/40 p-4 rounded-lg border border-slate-800/80"
                            >
                              <div>
                                <label className="text-[10px] font-bold text-indigo-400 uppercase block mb-1">Re-export Flight *</label>
                                <input
                                  type="text"
                                  required
                                  placeholder="LH761"
                                  value={bulkReexportFlight}
                                  onChange={(e) => setBulkReexportFlight(e.target.value.toUpperCase())}
                                  className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-2 px-3 rounded-lg"
                                />
                              </div>
                              <div>
                                <label className="text-[10px] font-bold text-indigo-400 uppercase block mb-1">Destination *</label>
                                <input
                                  type="text"
                                  required
                                  placeholder="FRA"
                                  value={bulkReexportDestination}
                                  onChange={(e) => setBulkReexportDestination(e.target.value.toUpperCase())}
                                  className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-2 px-3 rounded-lg"
                                />
                              </div>
                              <div className="col-span-2">
                                <label className="text-[10px] font-bold text-indigo-400 uppercase block mb-1">Re-export Date *</label>
                                <input
                                  type="date"
                                  required
                                  value={bulkReexportDate}
                                  onChange={(e) => setBulkReexportDate(e.target.value)}
                                  className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-2 px-3 rounded-lg"
                                />
                              </div>
                            </motion.div>
                          )}
                        </AnimatePresence>
                      </motion.div>
                    )}
                  </motion.div>
                )}

                {/* Step 2 for Cleared */}
                {bulkCustomsStatus === 'Cleared' && (
                  <motion.div
                    initial={{ opacity: 0, y: 10 }}
                    animate={{ opacity: 1, y: 0 }}
                    exit={{ opacity: 0, y: -10 }}
                    className="space-y-4"
                  >
                    <div className="space-y-1.5 bg-slate-950/40 p-4 rounded-lg border border-slate-800/80">
                      <label className="text-[10px] font-bold text-indigo-400 uppercase block tracking-wider">Step 2: Disposition *</label>
                      <select
                        required
                        value={bulkDisposition}
                        onChange={(e) => {
                          const disp = e.target.value as any;
                          setBulkDisposition(disp);
                          if (disp === 'Delivered') setBulkDeliveryDateTime(new Date().toISOString().slice(0, 16));
                          if (disp === 'Forwarded') setBulkDomesticForwarding('Air India');
                        }}
                        className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-sm py-2 px-3 rounded-lg cursor-pointer font-semibold focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 transition-all"
                      >
                        <option value="Pending">-- Select Disposition --</option>
                        <option value="Delivered">Delivery Agent</option>
                        <option value="Storage">LHG Office</option>
                        <option value="Forwarded">Domestic Forwarding</option>
                      </select>
                    </div>

                    {/* Dynamic Disposition Fields for Cleared */}
                    <AnimatePresence mode="wait">
                      {bulkDisposition === 'Delivered' && (
                        <motion.div
                          key="bulkDelivered"
                          initial={{ opacity: 0, height: 0 }}
                          animate={{ opacity: 1, height: 'auto' }}
                          exit={{ opacity: 0, height: 0 }}
                          className="space-y-4 bg-slate-950/40 p-4 rounded-lg border border-slate-800/80"
                        >
                          <div>
                            <label className="text-[10px] font-bold text-indigo-400 uppercase block mb-1">Delivery Agent *</label>
                            <select
                              required
                              value={bulkLocation}
                              onChange={(e) => setBulkLocation(e.target.value as any)}
                              className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-2 px-3 rounded-lg"
                            >
                              <option value="VVM">VVM</option>
                              <option value="Outlook">Outlook</option>
                              <option value="Advik">Advik</option>
                            </select>
                          </div>
                          <div>
                            <label className="text-[10px] font-bold text-indigo-400 uppercase block mb-1">Delivery Date & Time *</label>
                            <input
                              type="datetime-local"
                              required
                              value={bulkDeliveryDateTime}
                              onChange={(e) => setBulkDeliveryDateTime(e.target.value)}
                              className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-2 px-3 rounded-lg"
                            />
                          </div>
                        </motion.div>
                      )}

                      {bulkDisposition === 'Storage' && (
                        <motion.div
                          key="bulkStorage"
                          initial={{ opacity: 0, height: 0 }}
                          animate={{ opacity: 1, height: 'auto' }}
                          exit={{ opacity: 0, height: 0 }}
                          className="space-y-1.5 bg-slate-950/40 p-4 rounded-lg border border-slate-800/80"
                        >
                          <label className="text-[10px] font-bold text-indigo-400 uppercase block">Storage Remarks (Optional)</label>
                          <input
                            type="text"
                            placeholder="LHG Office storage details..."
                            value={bulkStorageRemarks}
                            onChange={(e) => setBulkStorageRemarks(e.target.value)}
                            className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-2 px-3 rounded-lg"
                          />
                        </motion.div>
                      )}

                      {bulkDisposition === 'Forwarded' && (
                        <motion.div
                          key="bulkForwarded"
                          initial={{ opacity: 0, height: 0 }}
                          animate={{ opacity: 1, height: 'auto' }}
                          exit={{ opacity: 0, height: 0 }}
                          className="grid grid-cols-2 gap-4 bg-slate-950/40 p-4 rounded-lg border border-slate-800/80"
                        >
                          <div className="col-span-2">
                            <label className="text-[10px] font-bold text-indigo-400 uppercase block mb-1">Forward Via *</label>
                            <select
                              required
                              value={bulkDomesticForwarding}
                              onChange={(e) => setBulkDomesticForwarding(e.target.value as any)}
                              className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-2 px-3 rounded-lg"
                            >
                              <option value="Air India">Air India</option>
                              <option value="SpiceJet">SpiceJet</option>
                              <option value="IndiGo">IndiGo</option>
                            </select>
                          </div>
                          <div>
                            <label className="text-[10px] font-bold text-indigo-400 uppercase block mb-1">Flight Number *</label>
                            <input
                              type="text"
                              required
                              placeholder="AI808"
                              value={bulkForwardingFlightNo}
                              onChange={(e) => setBulkForwardingFlightNo(e.target.value.toUpperCase())}
                              className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-2 px-3 rounded-lg"
                            />
                          </div>
                          <div>
                            <label className="text-[10px] font-bold text-indigo-400 uppercase block mb-1">Date *</label>
                            <input
                              type="date"
                              required
                              value={bulkForwardingDateNew}
                              onChange={(e) => setBulkForwardingDateNew(e.target.value)}
                              className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-2 px-3 rounded-lg"
                            />
                          </div>
                          <div className="col-span-2">
                            <label className="text-[10px] font-bold text-indigo-400 uppercase block mb-1">Destination *</label>
                            <input
                              type="text"
                              required
                              placeholder="DEL"
                              value={bulkForwardingDestinationNew}
                              onChange={(e) => setBulkForwardingDestinationNew(e.target.value.toUpperCase())}
                              className="w-full bg-slate-950 border border-slate-800 text-slate-200 text-xs py-2 px-3 rounded-lg"
                            />
                          </div>
                        </motion.div>
                      )}
                    </AnimatePresence>
                  </motion.div>
                )}
              </AnimatePresence>

              <div className="flex justify-end gap-2 pt-4 border-t border-slate-800">
                <button
                  type="button"
                  onClick={() => setShowBulkEdit(false)}
                  className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded text-xs transition"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  className="px-5 py-2 bg-indigo-600 hover:bg-indigo-500 text-white rounded text-xs font-semibold transition"
                >
                  Apply to {selectedIds.length} Records
                </button>
              </div>
            </form>

          </div>
        </div>
      )}

      {/* D. DELETE CONFIRMATION OVERLAY MODAL */}
      {deleteConfirm.show && (
        <div className="fixed inset-0 bg-black/80 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-800 w-full max-w-md rounded-xl shadow-2xl overflow-hidden p-6 space-y-4">
            <div className="flex items-center gap-3 text-red-500">
              <AlertOctagon className="w-6 h-6 animate-pulse" />
              <h3 className="font-bold text-slate-100 text-base">Confirm Permanent Deletion</h3>
            </div>
            
            <p className="text-slate-300 text-xs leading-relaxed">
              {deleteConfirm.type === 'bulk' 
                ? `Are you sure you want to permanently delete these ${selectedIds.length} selected baggage records? This action is irreversible.`
                : 'Are you sure you want to permanently delete this baggage record? This action is irreversible.'
              }
            </p>
            
            <div className="flex justify-end gap-2.5 pt-2 border-t border-slate-800/60">
              <button
                onClick={() => setDeleteConfirm({ show: false, type: 'single' })}
                className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded text-xs font-semibold transition"
              >
                Cancel
              </button>
              <button
                onClick={executeDeleteConfirmed}
                className="px-4 py-2 bg-red-600 hover:bg-red-500 text-white rounded text-xs font-bold transition shadow-lg shadow-red-600/10"
              >
                Confirm Delete
              </button>
            </div>
          </div>
        </div>
      )}

      {/* CAMERA BARCODE SCANNER MODAL */}
      <ScannerModal
        isOpen={isScannerOpen}
        onClose={() => setIsScannerOpen(false)}
        onScanSuccess={handleScanSuccess}
        title={scannerTargetField === 'bulk' ? 'Continuous Bulk Scanning (Bulk)' : 'Scan Baggage Barcode'}
        isContinuous={scannerTargetField === 'bulk'}
        continuousCount={continuousScannedTags.length}
        onFinishContinuous={handleFinishContinuous}
      />

      {/* DUPLICATE RESOLVER MODAL */}
      {showDuplicatesResolver && (
        <div className="fixed inset-0 bg-black/85 backdrop-blur-md z-50 flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-800 w-full max-w-2xl rounded-2xl shadow-2xl overflow-hidden flex flex-col animate-in fade-in zoom-in duration-200" style={{ maxHeight: '90vh' }}>
            {/* Header */}
            <div className="p-5 border-b border-slate-800 bg-slate-900/55 flex justify-between items-center shrink-0">
              <div className="flex items-center gap-2.5 text-amber-500">
                <AlertTriangle className="w-5 h-5 animate-pulse" />
                <div>
                  <h3 className="font-bold text-slate-100 text-sm">Duplicate Baggage Tags Detected</h3>
                  <p className="text-[10px] text-slate-400">Some scanned or entered tags conflict with existing records</p>
                </div>
              </div>
              <button 
                onClick={() => setShowDuplicatesResolver(false)}
                className="p-1 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-400 hover:text-slate-200 cursor-pointer transition"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            {/* Quick Batch Select Tools */}
            <div className="px-5 py-3.5 bg-slate-950 border-b border-slate-800/80 flex flex-col sm:flex-row sm:items-center justify-between gap-3 shrink-0">
              <span className="text-[10px] uppercase font-bold text-slate-400 tracking-wider">
                Batch Resolution Actions:
              </span>
              <div className="flex gap-2">
                <button
                  type="button"
                  onClick={() => {
                    setBulkDuplicatesList(prev => prev.map(d => ({ ...d, resolution: 'skip' })));
                  }}
                  className="px-3 py-1 bg-slate-900 hover:bg-slate-800 border border-slate-800 text-[10px] font-semibold text-slate-300 rounded cursor-pointer transition"
                >
                  Set All: Skip
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setBulkDuplicatesList(prev => prev.map(d => ({ ...d, resolution: 'replace' })));
                  }}
                  className="px-3 py-1 bg-indigo-950 hover:bg-indigo-900 border border-indigo-800/60 text-[10px] font-semibold text-indigo-300 rounded cursor-pointer transition"
                >
                  Set All: Replace
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setBulkDuplicatesList(prev => prev.map(d => ({ ...d, resolution: 'keep' })));
                  }}
                  className="px-3 py-1 bg-slate-900 hover:bg-slate-800 border border-slate-800 text-[10px] font-semibold text-emerald-400 rounded cursor-pointer transition"
                >
                  Set All: Keep Both
                </button>
              </div>
            </div>

            {/* Main scrollable area of duplicates list */}
            <div className="flex-1 overflow-y-auto p-5 space-y-3 bg-slate-900/40">
              {bulkDuplicatesList.map((item, index) => (
                <div 
                  key={item.tag}
                  className="p-4 bg-slate-950 rounded-xl border border-slate-850 flex flex-col md:flex-row md:items-center justify-between gap-4 transition hover:border-slate-800"
                >
                  {/* Left block: tag identifier info */}
                  <div className="space-y-1">
                    <div className="flex items-center gap-2">
                      <span className="text-xs font-mono font-bold text-slate-200 bg-slate-900 px-2.5 py-1 rounded border border-slate-800 tracking-widest">
                        {item.tag}
                      </span>
                      {item.isExistingInDb ? (
                        <span className="text-[9px] font-bold uppercase px-2 py-0.5 bg-red-950/60 text-red-400 border border-red-900/40 rounded">
                          DB Conflict
                        </span>
                      ) : (
                        <span className="text-[9px] font-bold uppercase px-2 py-0.5 bg-amber-950/60 text-amber-400 border border-amber-900/40 rounded">
                          Batch Dup
                        </span>
                      )}
                    </div>
                    
                    {item.existingRecord ? (
                      <p className="text-[10px] text-slate-500 leading-normal">
                        Passenger: <strong className="text-slate-300">{item.existingRecord.name}</strong> • PIR: <strong className="text-slate-300">{item.existingRecord.pir}</strong> • Flight: <strong className="text-slate-300">{item.existingRecord.flightNo}</strong>
                      </p>
                    ) : (
                      <p className="text-[10px] text-slate-500">
                        This tag is repeated multiple times in your current bulk operations list.
                      </p>
                    )}
                  </div>

                  {/* Right block: selector buttons */}
                  <div className="flex gap-1.5 self-start md:self-center">
                    {/* Skip button */}
                    <button
                      type="button"
                      onClick={() => {
                        setBulkDuplicatesList(prev => {
                          const next = [...prev];
                          next[index] = { ...next[index], resolution: 'skip' };
                          return next;
                        });
                      }}
                      className={`px-3 py-1.5 rounded text-[10px] font-bold border transition cursor-pointer ${
                        item.resolution === 'skip'
                          ? 'bg-slate-800 border-slate-700 text-slate-300'
                          : 'bg-slate-950 border-slate-900 text-slate-500 hover:border-slate-850 hover:text-slate-400'
                      }`}
                    >
                      Skip
                    </button>

                    {/* Replace / Update button */}
                    <button
                      type="button"
                      onClick={() => {
                        setBulkDuplicatesList(prev => {
                          const next = [...prev];
                          next[index] = { ...next[index], resolution: 'replace' };
                          return next;
                        });
                      }}
                      className={`px-3 py-1.5 rounded text-[10px] font-bold border transition cursor-pointer ${
                        item.resolution === 'replace'
                          ? 'bg-indigo-600 border-indigo-500 text-white shadow shadow-indigo-600/10'
                          : 'bg-slate-950 border-slate-900 text-slate-500 hover:border-slate-850 hover:text-slate-400'
                      }`}
                    >
                      Update Existing
                    </button>

                    {/* Keep / Create extra button */}
                    <button
                      type="button"
                      onClick={() => {
                        setBulkDuplicatesList(prev => {
                          const next = [...prev];
                          next[index] = { ...next[index], resolution: 'keep' };
                          return next;
                        });
                      }}
                      className={`px-3 py-1.5 rounded text-[10px] font-bold border transition cursor-pointer ${
                        item.resolution === 'keep'
                          ? 'bg-emerald-600 border-emerald-500 text-white shadow shadow-emerald-600/10'
                          : 'bg-slate-950 border-slate-900 text-slate-500 hover:border-slate-850 hover:text-slate-400'
                      }`}
                    >
                      Keep Both
                    </button>
                  </div>
                </div>
              ))}
            </div>

            {/* Footer with counts and process confirmation button */}
            <div className="p-4 bg-slate-900 border-t border-slate-800/80 flex flex-col sm:flex-row justify-between items-center gap-3 shrink-0">
              <div className="text-[10px] text-slate-400 text-center sm:text-left leading-normal font-mono">
                Skipping: <span className="text-slate-200 font-bold">{bulkDuplicatesList.filter(d => d.resolution === 'skip').length}</span> • 
                Updating: <span className="text-indigo-400 font-bold">{bulkDuplicatesList.filter(d => d.resolution === 'replace').length}</span> • 
                Keeping: <span className="text-emerald-400 font-bold">{bulkDuplicatesList.filter(d => d.resolution === 'keep').length}</span>
              </div>
              <div className="flex gap-2 w-full sm:w-auto">
                <button
                  type="button"
                  onClick={() => setShowDuplicatesResolver(false)}
                  className="flex-1 sm:flex-initial px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg text-xs font-bold cursor-pointer transition"
                >
                  Cancel Resolver
                </button>
                <button
                  type="button"
                  onClick={() => {
                    const tags = parseBulkInput(bulkTagsInput);
                    saveBulkRecords(tags, bulkDuplicatesList);
                  }}
                  className="flex-1 sm:flex-initial px-5 py-2 bg-gradient-to-r from-indigo-600 to-indigo-500 hover:from-indigo-500 hover:to-indigo-400 text-white rounded-lg text-xs font-bold cursor-pointer shadow-lg shadow-indigo-600/10 transition"
                >
                  Save &amp; Apply Resolutions
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

    </div>
  );
}
