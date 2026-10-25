---
layout: null
---
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Janak Positioning & Surveying - Stock Management</title>
    
    <!-- Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>
    
    <!-- Google Fonts -->
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;900&display=swap" rel="stylesheet">
    
    <!-- Phosphor Icons -->
    <script src="https://unpkg.com/@phosphor-icons/web"></script>

    <!-- SheetJS for Excel (.xlsx) Export -->
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    
    <style>
        body {
            font-family: 'Inter', sans-serif;
            background-color: #f8fafc;
            margin: 0;
            padding: 0;
            min-height: 100vh;
            overflow-y: auto; /* Fixed: Allows full page natural scrolling */
            overflow-x: hidden;
        }
        
        /* Custom Scrollbar */
        ::-webkit-scrollbar { width: 8px; height: 8px; }
        ::-webkit-scrollbar-track { background: #f1f5f9; border-radius: 4px; }
        ::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 4px; }
        ::-webkit-scrollbar-thumb:hover { background: #94a3b8; }

        /* Print Styles */
        @media print {
            body { background-color: white; overflow: auto; height: auto; }
            #app-container, #login-screen, #toast-container, .no-print, #edit-modal, #camera-modal, #approval-details-modal, #gatepass-preview-modal, #csv-import-modal, #track-modal, #dispatch-details-modal, #manage-models-modal, #drive-sync-modal, #model-pin-modal, #user-edit-modal, #forgot-password-modal {
                display: none !important;
            }
            #gatepass-print-view {
                display: block !important;
                position: absolute;
                top: 0;
                left: 0;
                width: 100%;
                margin: 0;
                padding: 20px;
                background-color: white;
            }
        }
        
        @media screen {
            #gatepass-print-view {
                display: none;
            }
        }
    </style>
</head>
<body class="text-slate-800 antialiased flex flex-col bg-slate-50 min-h-screen">

    <!-- Login Screen -->
    <div id="login-screen" class="fixed inset-0 z-[100] flex items-center justify-center bg-slate-900 px-4">
        <div class="bg-white p-8 rounded-2xl shadow-2xl max-w-sm w-full">
            <div class="flex justify-center mb-2">
                <h2 class="text-3xl font-black text-red-600 text-center uppercase tracking-wide">Janak Positioning</h2>
            </div>
            <h3 class="text-sm font-bold text-center text-slate-800 mb-2 mt-2">Stock Management App</h3>
            <p class="text-center text-slate-500 text-xs mb-6">Sign in to your account</p>
            
            <div id="login-error" class="hidden mb-4 p-3 bg-rose-50 text-rose-600 text-sm rounded-lg border border-rose-200 text-center font-medium">
                Invalid username or password
            </div>

            <!-- Quick Role Select Buttons -->
            <div class="mb-5">
                <p class="text-xs font-semibold text-slate-400 uppercase tracking-wider text-center mb-2.5">Select Role / Quick Login</p>
                <div class="grid grid-cols-2 gap-2.5">
                    <button type="button" id="role-btn-admin" onclick="selectLoginRole('admin')" class="py-2.5 px-3 bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-200 text-xs font-bold rounded-xl transition flex items-center justify-center gap-1.5 shadow-sm">
                        <i class="ph-fill ph-shield-check text-slate-500 text-base" id="role-icon-admin"></i> Admin
                    </button>
                    <button type="button" id="role-btn-manager" onclick="selectLoginRole('manager')" class="selectLoginRole('manager')" onclick="selectLoginRole('manager')" class="py-2.5 px-3 bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-200 text-xs font-bold rounded-xl transition flex items-center justify-center gap-1.5 shadow-sm">
                        <i class="ph-fill ph-user-circle text-slate-500 text-base" id="role-icon-manager"></i> Store Manager
                    </button>
                </div>
            </div>

            <form id="form-login" onsubmit="handleLogin(event)" class="space-y-4">
                <div>
                    <label class="block text-sm font-semibold text-slate-700 mb-1">Username</label>
                    <input type="text" id="login-user" required class="w-full px-4 py-3 border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none">
                </div>
                <div>
                    <label class="block text-sm font-semibold text-slate-700 mb-1">Password</label>
                    <input type="password" id="login-pass" required class="w-full px-4 py-3 border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none">
                </div>
                
                <!-- Forgot Password Trigger -->
                <div class="flex items-center justify-end text-xs">
                    <button type="button" onclick="openForgotPasswordModal()" class="text-blue-600 hover:underline font-semibold">Forgot Password?</button>
                </div>

                <button type="submit" class="w-full bg-blue-600 text-white font-bold py-3 px-4 rounded-xl hover:bg-blue-700 transition shadow-md">
                    Login to System
                </button>
            </form>
        </div>
    </div>

    <!-- App Container (Fixed: Removed h-screen and overflow-hidden constraints to enable full natural page scroll) -->
    <div id="app-container" class="hidden w-full flex-1 flex-col md:flex-row min-h-screen">
        
        <!-- Sidebar -->
        <aside class="w-full md:w-64 bg-slate-900 text-white flex flex-col shadow-xl z-20 shrink-0">
            <div class="p-6 border-b border-slate-800 flex flex-col justify-center">
                <h1 class="font-black text-2xl leading-tight text-red-600 uppercase tracking-wider">Janak Positioning</h1>
                <p class="text-slate-400 font-medium text-xs mt-1 uppercase tracking-widest">Surveying Systems</p>
            </div>
            
            <nav class="flex-1 py-4 md:py-6 px-3 flex md:flex-col gap-2 overflow-x-auto md:overflow-visible">
                <button type="button" onclick="switchTab('dashboard')" id="nav-dashboard" class="nav-btn whitespace-nowrap md:w-full flex items-center gap-3 px-4 py-3 rounded-xl bg-blue-600 text-white font-medium transition-colors text-left shrink-0">
                    <i class="ph ph-squares-four text-xl"></i>
                    <span>Dashboard</span>
                </button>
                <button type="button" onclick="switchTab('material-in')" id="nav-material-in" class="nav-btn whitespace-nowrap md:w-full flex items-center gap-3 px-4 py-3 rounded-xl text-slate-400 hover:bg-slate-800 hover:text-white font-medium transition-colors text-left shrink-0">
                    <i class="ph ph-box-arrow-down text-xl"></i>
                    <span>Material In</span>
                </button>
                <button type="button" onclick="switchTab('material-out')" id="nav-material-out" class="nav-btn whitespace-nowrap md:w-full flex items-center gap-3 px-4 py-3 rounded-xl text-slate-400 hover:bg-slate-800 hover:text-white font-medium transition-colors text-left shrink-0">
                    <i class="ph ph-truck text-xl"></i>
                    <span>Material Out</span>
                </button>
                <button type="button" onclick="switchTab('dispatch-history')" id="nav-dispatch-history" class="nav-btn whitespace-nowrap md:w-full flex items-center gap-3 px-4 py-3 rounded-xl text-slate-400 hover:bg-slate-800 hover:text-white font-medium transition-colors text-left shrink-0">
                    <i class="ph ph-clock-counter-clockwise text-xl"></i>
                    <span>Dispatch History</span>
                </button>
                <button type="button" onclick="switchTab('approvals')" id="nav-approvals" class="nav-btn hidden whitespace-nowrap md:w-full flex items-center gap-3 px-4 py-3 rounded-xl text-slate-400 hover:bg-slate-800 hover:text-white font-medium transition-colors text-left shrink-0">
                    <i class="ph ph-check-square text-xl"></i>
                    <span>Approvals</span>
                </button>
                <button type="button" onclick="switchTab('user-management')" id="nav-user-management" class="nav-btn hidden whitespace-nowrap md:w-full flex items-center gap-3 px-4 py-3 rounded-xl text-slate-400 hover:bg-slate-800 hover:text-white font-medium transition-colors text-left shrink-0">
                    <i class="ph ph-users-three text-xl"></i>
                    <span>User Management</span>
                </button>
                <button type="button" onclick="requestManageModelsAccess()" id="nav-manage-models" class="hidden whitespace-nowrap md:w-full flex items-center gap-3 px-4 py-3 rounded-xl text-indigo-300 hover:bg-slate-800 hover:text-white font-medium transition-colors text-left shrink-0">
                    <i class="ph ph-gear-six text-xl"></i>
                    <span>Manage Models</span>
                </button>
            </nav>
            
            <div class="p-4 border-t border-slate-800 hidden md:block">
                <div class="flex items-center justify-between gap-3">
                    <div class="flex items-center gap-3">
                        <div class="w-8 h-8 rounded-full bg-slate-700 flex items-center justify-center">
                            <i class="ph ph-user text-slate-300"></i>
                        </div>
                        <div>
                            <p class="text-xs text-slate-400">Logged in as</p>
                            <p id="user-role-display" class="text-sm font-semibold text-emerald-400">Admin</p>
                        </div>
                    </div>
                    <button type="button" onclick="logout()" class="p-2 text-slate-400 hover:text-rose-400 hover:bg-slate-800 rounded-lg transition" title="Logout">
                        <i class="ph ph-sign-out text-lg"></i>
                    </button>
                </div>
            </div>
        </aside>

        <!-- Main Content Area -->
        <main class="flex-1 flex flex-col relative bg-slate-50 min-w-0">
            
            <!-- Header -->
            <header class="h-16 bg-white border-b border-slate-200 flex items-center justify-between px-4 md:px-8 z-10 shrink-0 sticky top-0">
                <div class="flex items-center gap-3">
                    <h2 class="text-xl font-bold text-slate-800" id="header-title">Inventory Dashboard</h2>
                    <button type="button" onclick="openDriveSyncModal()" id="drive-status-badge" class="px-2.5 py-1 rounded-full text-[11px] font-bold border flex items-center gap-1.5 transition bg-slate-100 text-slate-600 border-slate-200" title="Click to configure Google Drive Sync">
                        <span id="drive-status-dot" class="w-2 h-2 rounded-full bg-slate-400"></span>
                        <span id="drive-status-text">Drive: Not Linked</span>
                    </button>
                </div>
                <div class="flex items-center gap-2 flex-wrap">
                    <button type="button" id="admin-manage-models-btn" onclick="requestManageModelsAccess()" class="hidden px-3 py-1.5 bg-indigo-50 text-indigo-700 rounded-full text-xs font-semibold border border-indigo-200 items-center gap-1 hover:bg-indigo-100 transition">
                        <i class="ph ph-gear-six"></i> <span>Manage Models</span>
                    </button>
                    <button type="button" onclick="openCsvModal()" class="px-3 py-1.5 bg-emerald-50 text-emerald-700 rounded-full text-xs font-semibold border border-emerald-200 flex items-center gap-1 hover:bg-emerald-100 transition" title="Import CSV">
                        <i class="ph ph-file-arrow-up"></i> <span class="hidden sm:inline">Import CSV</span>
                    </button>
                    <button type="button" onclick="exportToExcel()" class="px-3 py-1.5 bg-teal-50 text-teal-700 rounded-full text-xs font-semibold border border-teal-200 flex items-center gap-1 hover:bg-teal-100 transition" title="Download Excel (.xlsx)">
                        <i class="ph ph-file-xls"></i> <span class="hidden sm:inline">Excel Backup</span>
                    </button>
                    <button type="button" onclick="openDriveSyncModal()" class="px-3 py-1.5 bg-amber-50 text-amber-700 rounded-full text-xs font-semibold border border-amber-200 flex items-center gap-1 hover:bg-amber-100 transition" title="Google Drive Auto-Sync">
                        <i class="ph ph-cloud-arrow-up"></i> <span class="hidden sm:inline">Drive Sync</span>
                    </button>
                    <button type="button" onclick="logout()" class="md:hidden p-2 text-slate-500 hover:text-rose-500 rounded-lg transition" title="Logout">
                        <i class="ph ph-sign-out text-xl"></i>
                    </button>
                </div>
            </header>

            <!-- Scrollable Viewport Area (Fixed: Removed inner height locks) -->
            <div id="main-scroll-area" class="flex-1 p-4 md:p-8 relative">
                
                <!-- Tab: Dashboard -->
                <div id="tab-dashboard" class="tab-content max-w-6xl mx-auto block space-y-6 pb-12">
                    <div class="grid grid-cols-1 sm:grid-cols-3 gap-4 md:gap-6">
                        <div class="bg-white p-6 rounded-2xl shadow-sm border border-slate-200">
                            <div class="flex justify-between items-start">
                                <div>
                                    <p class="text-slate-500 text-sm font-medium mb-1">Total Assets</p>
                                    <h3 class="text-3xl font-bold text-slate-900" id="stat-total">0</h3>
                                </div>
                                <div class="w-10 h-10 rounded-full bg-blue-50 flex items-center justify-center text-blue-600"><i class="ph ph-files text-xl"></i></div>
                            </div>
                        </div>
                        <div class="bg-white p-6 rounded-2xl shadow-sm border border-slate-200">
                            <div class="flex justify-between items-start">
                                <div>
                                    <p class="text-slate-500 text-sm font-medium mb-1">Currently In Stock</p>
                                    <h3 class="text-3xl font-bold text-emerald-600" id="stat-in">0</h3>
                                </div>
                                <div class="w-10 h-10 rounded-full bg-emerald-50 flex items-center justify-center text-emerald-600"><i class="ph ph-package text-xl"></i></div>
                            </div>
                        </div>
                        <div class="bg-white p-6 rounded-2xl shadow-sm border border-slate-200">
                            <div class="flex justify-between items-start">
                                <div>
                                    <p class="text-slate-500 text-sm font-medium mb-1">Dispatched (Out)</p>
                                    <h3 class="text-3xl font-bold text-amber-600" id="stat-out">0</h3>
                                </div>
                                <div class="w-10 h-10 rounded-full bg-amber-50 flex items-center justify-center text-amber-600"><i class="ph ph-truck text-xl"></i></div>
                            </div>
                        </div>
                    </div>

                    <!-- Category Breakdown Section -->
                    <div class="bg-white p-6 rounded-2xl shadow-sm border border-slate-200">
                        <div class="flex items-center justify-between mb-4">
                            <h3 class="font-bold text-slate-800 text-base flex items-center gap-2">
                                <i class="ph ph-chart-bar text-blue-600 text-xl"></i> Category Stock Summary
                            </h3>
                            <span class="text-xs text-slate-500">Live In-Stock Counts & Model Breakdown</span>
                        </div>
                        <div id="category-summary-grid" class="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-4">
                            <!-- Populated dynamically -->
                        </div>
                    </div>

                    <!-- Low Stock Dashboard Widget -->
                    <div id="low-stock-widget-container" class="hidden bg-rose-50 border border-rose-200 rounded-2xl p-5 shadow-sm">
                        <div class="flex items-center gap-3 mb-3">
                            <div class="w-9 h-9 rounded-xl bg-rose-600 text-white flex items-center justify-center shrink-0">
                                <i class="ph-fill ph-warning text-xl"></i>
                            </div>
                            <div>
                                <h3 class="font-bold text-rose-900 text-base">Low Stock Warnings (Threshold: &le; 5 Units)</h3>
                                <p class="text-xs text-rose-700">The following models require immediate restocking.</p>
                            </div>
                        </div>
                        <div id="low-stock-list" class="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3">
                            <!-- Populated dynamically -->
                        </div>
                    </div>

                    <div class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden flex flex-col">
                        <div class="px-6 py-4 border-b border-slate-200 bg-slate-50 flex flex-col sm:flex-row justify-between items-center gap-4">
                            <h3 class="font-bold text-slate-800">Master Inventory List</h3>
                            <div class="relative w-full sm:w-64">
                                <i class="ph ph-magnifying-glass absolute left-3 top-1/2 transform -translate-y-1/2 text-slate-400"></i>
                                <input type="text" id="search-input" placeholder="Search item, source or S/N..." class="w-full pl-9 pr-4 py-2 text-sm border border-slate-300 rounded-lg focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500" onkeyup="renderTable()">
                            </div>
                        </div>
                        <div class="overflow-x-auto">
                            <table class="w-full text-left border-collapse min-w-[900px]">
                                <thead>
                                    <tr class="bg-white text-slate-500 text-xs uppercase tracking-wider border-b border-slate-200">
                                        <th class="px-6 py-4 font-semibold">Category & Model</th>
                                        <th class="px-6 py-4 font-semibold">Serial Number</th>
                                        <th class="px-6 py-4 font-semibold">Status</th>
                                        <th class="px-6 py-4 font-semibold">Activity Info</th>
                                        <th class="px-6 py-4 font-semibold text-right">Actions</th>
                                    </tr>
                                </thead>
                                <tbody id="inventory-table-body" class="divide-y divide-slate-100 bg-white">
                                    <!-- Dynamic rows -->
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>

                <!-- Tab: Material In -->
                <div id="tab-material-in" class="tab-content max-w-2xl mx-auto hidden pb-12">
                    <div class="bg-white rounded-2xl shadow-sm border border-slate-200 p-6 md:p-8">
                        <div class="flex items-center justify-between mb-6">
                            <div class="flex items-center gap-3">
                                <div class="w-12 h-12 bg-blue-50 text-blue-600 rounded-xl flex items-center justify-center">
                                    <i class="ph ph-box-arrow-down text-2xl"></i>
                                </div>
                                <div>
                                    <h3 class="text-xl font-bold text-slate-800">Receive Material</h3>
                                    <p class="text-sm text-slate-500">Log new equipment into the local database.</p>
                                </div>
                            </div>
                            <button type="button" id="in-tab-manage-btn" onclick="requestManageModelsAccess()" class="hidden px-3 py-2 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 rounded-xl text-xs font-bold transition items-center gap-1.5 border border-indigo-200">
                                <i class="ph ph-plus-circle text-base"></i> Add / Edit Models
                            </button>
                        </div>

                        <form id="form-material-in" class="space-y-4">
                            <div>
                                <label class="block text-sm font-semibold text-slate-700 mb-1">Equipment Model / Name <span class="text-rose-500">*</span></label>
                                <select id="in-name" required class="w-full px-4 py-3 border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none transition bg-white">
                                    <!-- Populated dynamically via JS -->
                                </select>
                            </div>
                            
                            <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                                <div>
                                    <label class="block text-sm font-semibold text-slate-700 mb-1">Serial Number (S/N) <span class="text-rose-500">*</span></label>
                                    <input type="text" id="in-sn" required placeholder="e.g. SN-104958" 
                                        class="w-full px-4 py-3 border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none transition uppercase">
                                </div>
                                <div>
                                    <label class="block text-sm font-semibold text-slate-700 mb-1">Received From (Source) <span class="text-rose-500">*</span></label>
                                    <input type="text" id="in-source" required placeholder="e.g. Factory / Vendor / Head Office" 
                                        class="w-full px-4 py-3 border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none transition">
                                </div>
                            </div>

                            <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                                <div>
                                    <label class="block text-sm font-semibold text-slate-700 mb-1">Date Received <span class="text-rose-500">*</span></label>
                                    <input type="date" id="in-date" required 
                                        class="w-full px-4 py-3 border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none transition">
                                </div>
                                <div>
                                    <label class="block text-sm font-semibold text-slate-700 mb-1">Accessories Included</label>
                                    <input type="text" id="in-accessories" placeholder="e.g. 2 Batteries, Charger, Hard Case" 
                                        class="w-full px-4 py-3 border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none transition">
                                </div>
                            </div>

                            <button type="submit" class="w-full mt-4 bg-blue-600 text-white font-semibold py-3.5 px-4 rounded-xl hover:bg-blue-700 transition duration-200 flex justify-center items-center gap-2 shadow-sm">
                                <i class="ph ph-plus-circle text-lg"></i> Add to Inventory
                            </button>
                        </form>
                    </div>
                </div>

                <!-- Tab: Material Out -->
                <div id="tab-material-out" class="tab-content max-w-2xl mx-auto hidden pb-12">
                    <div class="bg-white rounded-2xl shadow-sm border border-slate-200 p-6 md:p-8">
                        <div class="flex items-center gap-3 mb-6">
                            <div class="w-12 h-12 bg-amber-50 text-amber-600 rounded-xl flex items-center justify-center">
                                <i class="ph ph-truck text-2xl"></i>
                            </div>
                            <div>
                                <h3 class="text-xl font-bold text-slate-800">Dispatch Material</h3>
                                <p class="text-sm text-slate-500">Provide recipient details and choose dispatch mode.</p>
                            </div>
                        </div>

                        <form id="form-material-out" class="space-y-4">
                            <div>
                                <label class="block text-sm font-semibold text-slate-700 mb-1">Select Equipment (Currently In Stock) <span class="text-rose-500">*</span></label>
                                <select id="out-item" required class="w-full px-4 py-3 border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none transition bg-white text-slate-700">
                                    <option value="" disabled selected>Select an item...</option>
                                </select>
                            </div>

                            <!-- Dispatch Mode Selection -->
                            <div>
                                <label class="block text-sm font-semibold text-slate-700 mb-2">Dispatch Mode <span class="text-rose-500">*</span></label>
                                <div class="grid grid-cols-2 gap-4">
                                    <label class="flex items-center gap-3 p-3 border border-slate-300 rounded-xl cursor-pointer hover:bg-slate-50 transition">
                                        <input type="radio" name="dispatch-mode" value="Direct" checked onchange="toggleDispatchMode()" class="text-blue-600 focus:ring-blue-500">
                                        <div>
                                            <span class="font-semibold text-sm block text-slate-800">Direct / Hand Delivery</span>
                                            <span class="text-xs text-slate-500">Standard site challan</span>
                                        </div>
                                    </label>
                                    <label class="flex items-center gap-3 p-3 border border-slate-300 rounded-xl cursor-pointer hover:bg-slate-50 transition">
                                        <input type="radio" name="dispatch-mode" value="Courier" onchange="toggleDispatchMode()" class="text-blue-600 focus:ring-blue-500">
                                        <div>
                                            <span class="font-semibold text-sm block text-slate-800">Courier / Logistics</span>
                                            <span class="text-xs text-slate-500">Shipped via service</span>
                                        </div>
                                    </label>
                                </div>
                            </div>
                            
                            <!-- Courier Specific Fields -->
                            <div id="courier-fields" class="hidden space-y-4 p-4 bg-slate-50 rounded-xl border border-slate-200">
                                <h4 class="text-xs font-bold uppercase tracking-wider text-slate-500 flex items-center gap-1.5">
                                    <i class="ph ph-package text-blue-600"></i> Courier Logistics Details
                                </h4>
                                <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                                    <div>
                                        <label class="block text-sm font-semibold text-slate-700 mb-1">Courier Partner Name <span class="text-rose-500">*</span></label>
                                        <input type="text" id="out-courier-name" placeholder="e.g. Blue Dart / DTDC / Professional" 
                                            class="w-full px-4 py-2.5 border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none transition bg-white text-sm">
                                    </div>
                                    <div>
                                        <label class="block text-sm font-semibold text-slate-700 mb-1">AWB / Tracking Number <span class="text-slate-400 font-normal">(Optional)</span></label>
                                        <input type="text" id="out-awb" placeholder="e.g. AW-984729104" 
                                            class="w-full px-4 py-2.5 border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none transition bg-white uppercase text-sm">
                                    </div>
                                </div>
                            </div>
                            
                            <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                                <div>
                                    <label class="block text-sm font-semibold text-slate-700 mb-1">Company / Client Name <span class="text-rose-500">*</span></label>
                                    <input type="text" id="out-company" required placeholder="e.g. Acme Surveyors Pvt Ltd" 
                                        class="w-full px-4 py-3 border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none transition">
                                </div>
                                <div>
                                    <label class="block text-sm font-semibold text-slate-700 mb-1">Challan / Invoice Number <span class="text-slate-400 font-normal">(Optional)</span></label>
                                    <input type="text" id="out-challan" placeholder="e.g. INV-2026-089" 
                                        class="w-full px-4 py-3 border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none transition">
                                </div>
                            </div>

                            <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                                <div>
                                    <label class="block text-sm font-semibold text-slate-700 mb-1">Contact Phone Number <span class="text-rose-500">*</span></label>
                                    <input type="text" id="out-phone" required placeholder="e.g. +91 9876543210" 
                                        class="w-full px-4 py-3 border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none transition">
                                </div>
                                <div>
                                    <label class="block text-sm font-semibold text-slate-700 mb-1">Site / Delivery Location <span class="text-rose-500">*</span></label>
                                    <input type="text" id="out-to" required placeholder="e.g. Sector 62, Noida Site A" 
                                        class="w-full px-4 py-3 border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none transition">
                                </div>
                            </div>

                            <div>
                                <label class="block text-sm font-semibold text-slate-700 mb-1">Full Details / Remarks / Accessories</label>
                                <textarea id="out-remarks" rows="2" placeholder="e.g. Issued with 2 batteries, charger, hard case, calibration certificate." 
                                    class="w-full px-4 py-3 border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none transition"></textarea>
                            </div>

                            <button type="submit" class="w-full mt-4 bg-slate-900 text-white font-semibold py-3.5 px-4 rounded-xl hover:bg-slate-800 transition duration-200 flex justify-center items-center gap-2 shadow-sm">
                                <i class="ph ph-printer text-lg"></i> Request Dispatch
                            </button>
                        </form>
                    </div>
                </div>

                <!-- Tab: Dispatch History -->
                <div id="tab-dispatch-history" class="tab-content max-w-6xl mx-auto hidden pb-12">
                    <div class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden flex flex-col">
                        <div class="px-6 py-4 border-b border-slate-200 bg-slate-50 flex flex-col md:flex-row justify-between items-center gap-4">
                            <div>
                                <h3 class="font-bold text-slate-800 text-lg">Dispatch History Log</h3>
                                <p class="text-xs text-slate-500">Filter and audit all past dispatched equipment date-wise.</p>
                            </div>
                            <div class="flex flex-wrap items-center gap-3 w-full md:w-auto">
                                <div class="flex items-center gap-2 text-xs">
                                    <span class="text-slate-500 font-medium">From:</span>
                                    <input type="date" id="filter-date-from" onchange="renderDispatchHistory()" class="px-3 py-1.5 border border-slate-300 rounded-lg text-xs bg-white outline-none focus:border-blue-500">
                                </div>
                                <div class="flex items-center gap-2 text-xs">
                                    <span class="text-slate-500 font-medium">To:</span>
                                    <input type="date" id="filter-date-to" onchange="renderDispatchHistory()" class="px-3 py-1.5 border border-slate-300 rounded-lg text-xs bg-white outline-none focus:border-blue-500">
                                </div>
                                <button type="button" onclick="resetDispatchFilters()" class="px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg text-xs font-semibold transition">Reset</button>
                            </div>
                        </div>
                        <div class="overflow-x-auto">
                            <table class="w-full text-left border-collapse min-w-[900px]">
                                <thead>
                                    <tr class="bg-white text-slate-500 text-xs uppercase tracking-wider border-b border-slate-200">
                                        <th class="px-6 py-4 font-semibold">Equipment & S/N</th>
                                        <th class="px-6 py-4 font-semibold">Client & Challan</th>
                                        <th class="px-6 py-4 font-semibold">Dispatch Mode & Info</th>
                                        <th class="px-6 py-4 font-semibold">Date Out</th>
                                        <th class="px-6 py-4 font-semibold text-right">Actions</th>
                                    </tr>
                                </thead>
                                <tbody id="dispatch-history-table-body" class="divide-y divide-slate-100 bg-white">
                                    <!-- Dynamic rows -->
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>

                <!-- Tab: Approvals -->
                <div id="tab-approvals" class="tab-content max-w-4xl mx-auto hidden pb-12">
                    <div class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden flex flex-col">
                        <div class="px-6 py-4 border-b border-slate-200 bg-slate-50 flex items-center gap-3">
                            <i class="ph-fill ph-shield-check text-blue-600 text-xl"></i>
                            <h3 class="font-bold text-slate-800">Pending Dispatch Approvals</h3>
                        </div>
                        <div class="overflow-x-auto">
                            <table class="w-full text-left border-collapse">
                                <thead>
                                    <tr class="bg-white text-slate-500 text-xs uppercase tracking-wider border-b border-slate-200">
                                        <th class="px-6 py-4 font-semibold">Equipment & S/N</th>
                                        <th class="px-6 py-4 font-semibold">Client & Mode</th>
                                        <th class="px-6 py-4 font-semibold text-right">Actions</th>
                                    </tr>
                                </thead>
                                <tbody id="approvals-table-body" class="divide-y divide-slate-100 bg-white">
                                    <!-- Dynamic rows -->
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>

                <!-- Tab: User Management -->
                <div id="tab-user-management" class="tab-content max-w-5xl mx-auto hidden space-y-6 pb-12">
                    <div class="bg-white rounded-2xl shadow-sm border border-slate-200 p-6 md:p-8">
                        <div class="flex items-center gap-3 mb-6">
                            <div class="w-12 h-12 bg-indigo-50 text-indigo-600 rounded-xl flex items-center justify-center">
                                <i class="ph ph-user-plus text-2xl"></i>
                            </div>
                            <div>
                                <h3 class="text-xl font-bold text-slate-800">Create New User Account</h3>
                                <p class="text-sm text-slate-500">Add an administrator or store manager credential.</p>
                            </div>
                        </div>

                        <form id="form-create-user" onsubmit="handleCreateUser(event)" class="grid grid-cols-1 sm:grid-cols-3 gap-4">
                            <div>
                                <label class="block text-sm font-semibold text-slate-700 mb-1">Username <span class="text-rose-500">*</span></label>
                                <input type="text" id="new-user-name" required placeholder="e.g. john_manager" class="w-full px-4 py-2.5 border border-slate-300 rounded-xl text-sm outline-none focus:ring-2 focus:ring-blue-500 bg-white">
                            </div>
                            <div>
                                <label class="block text-sm font-semibold text-slate-700 mb-1">Password <span class="text-rose-500">*</span></label>
                                <input type="password" id="new-user-pass" required placeholder="••••••••" class="w-full px-4 py-2.5 border border-slate-300 rounded-xl text-sm outline-none focus:ring-2 focus:ring-blue-500 bg-white">
                            </div>
                            <div>
                                <label class="block text-sm font-semibold text-slate-700 mb-1">Role <span class="text-rose-500">*</span></label>
                                <div class="flex gap-2">
                                    <select id="new-user-role" required class="flex-1 px-4 py-2.5 border border-slate-300 rounded-xl text-sm outline-none focus:ring-2 focus:ring-blue-500 bg-white">
                                        <option value="manager">Store Manager</option>
                                        <option value="admin">Admin</option>
                                    </select>
                                    <button type="submit" class="px-5 py-2.5 bg-indigo-600 hover:bg-indigo-700 text-white font-semibold rounded-xl text-sm transition shadow-sm shrink-0">
                                        Create User
                                    </button>
                                </div>
                            </div>
                        </form>
                    </div>

                    <!-- User Accounts Table -->
                    <div class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden flex flex-col">
                        <div class="px-6 py-4 border-b border-slate-200 bg-slate-50 flex items-center justify-between">
                            <div>
                                <h3 class="font-bold text-slate-800 text-base">Registered System Accounts</h3>
                                <p class="text-xs text-slate-500">Manage usernames, roles, or remove accounts.</p>
                            </div>
                            <span id="user-count-badge" class="px-2.5 py-1 bg-slate-200 text-slate-700 rounded-full text-xs font-bold">0 Users</span>
                        </div>
                        <div class="overflow-x-auto">
                            <table class="w-full text-left border-collapse">
                                <thead>
                                    <tr class="bg-white text-slate-500 text-xs uppercase tracking-wider border-b border-slate-200">
                                        <th class="px-6 py-4 font-semibold">Username</th>
                                        <th class="px-6 py-4 font-semibold">Role</th>
                                        <th class="px-6 py-4 font-semibold">Created Date</th>
                                        <th class="px-6 py-4 font-semibold text-right">Actions</th>
                                    </tr>
                                </thead>
                                <tbody id="users-table-body" class="divide-y divide-slate-100 bg-white">
                                    <!-- Dynamic rows -->
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>

            </div>
        </main>
    </div>

    <!-- Password Prompt Modal for Manage Models (Passcode: 8317) -->
    <div id="model-pin-modal" class="hidden fixed inset-0 z-[110] bg-slate-900/70 backdrop-blur-sm flex items-center justify-center p-4">
        <div class="bg-white p-6 rounded-2xl shadow-2xl max-w-xs w-full text-center space-y-4">
            <div class="w-12 h-12 bg-indigo-50 text-indigo-600 rounded-xl flex items-center justify-center mx-auto text-2xl">
                <i class="ph-fill ph-lock-key"></i>
            </div>
            <div>
                <h3 class="font-bold text-slate-800 text-lg">Enter Access PIN</h3>
                <p class="text-xs text-slate-500 mt-0.5">Please enter PIN to manage models (8317)</p>
            </div>
            <div id="pin-error" class="hidden text-rose-600 text-xs font-semibold">Incorrect PIN. Try again (8317).</div>
            <form onsubmit="verifyModelPin(event)" class="space-y-3">
                <input type="password" id="model-pin-input" autofocus maxlength="4" placeholder="••••" required class="w-full px-4 py-3 text-center text-xl tracking-widest border border-slate-300 rounded-xl focus:ring-2 focus:ring-indigo-500 outline-none font-mono">
                <div class="flex gap-2">
                    <button type="button" onclick="closeModelPinModal()" class="flex-1 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-semibold rounded-xl transition">Cancel</button>
                    <button type="submit" class="flex-1 py-2.5 bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-bold rounded-xl transition">Unlock</button>
                </div>
            </form>
        </div>
    </div>

    <!-- Forgot Password Recovery Modal (Passcode: 8317) -->
    <div id="forgot-password-modal" class="hidden fixed inset-0 z-[120] bg-slate-900/70 backdrop-blur-sm flex items-center justify-center p-4">
        <div class="bg-white p-6 rounded-2xl shadow-2xl max-w-sm w-full space-y-4">
            <div class="w-12 h-12 bg-amber-50 text-amber-600 rounded-xl flex items-center justify-center mx-auto text-2xl">
                <i class="ph-fill ph-key"></i>
            </div>
            <div>
                <h3 class="font-bold text-slate-800 text-lg text-center">Reset Admin Password</h3>
                <p class="text-xs text-slate-500 text-center mt-0.5">Enter Master Recovery PIN (8317) to reset credentials</p>
            </div>
            <div id="recovery-error" class="hidden text-rose-600 text-xs font-semibold text-center">Incorrect Recovery PIN.</div>
            <form onsubmit="handlePasswordRecovery(event)" class="space-y-3">
                <div>
                    <label class="block text-xs font-semibold text-slate-700 mb-1">Master Recovery PIN</label>
                    <input type="password" id="recovery-pin" required maxlength="4" placeholder="••••" class="w-full px-4 py-2.5 text-center text-lg tracking-widest border border-slate-300 rounded-xl focus:ring-2 focus:ring-amber-500 outline-none font-mono">
                </div>
                <div>
                    <label class="block text-xs font-semibold text-slate-700 mb-1">New Admin Username</label>
                    <input type="text" id="recovery-new-user" required value="admin" class="w-full px-3 py-2 border border-slate-300 rounded-xl text-sm outline-none focus:ring-2 focus:ring-amber-500">
                </div>
                <div>
                    <label class="block text-xs font-semibold text-slate-700 mb-1">New Admin Password</label>
                    <input type="password" id="recovery-new-pass" required placeholder="Enter new password" class="w-full px-3 py-2 border border-slate-300 rounded-xl text-sm outline-none focus:ring-2 focus:ring-amber-500">
                </div>
                <div class="flex gap-2 pt-2">
                    <button type="button" onclick="closeForgotPasswordModal()" class="flex-1 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-semibold rounded-xl transition">Cancel</button>
                    <button type="submit" class="flex-1 py-2 bg-amber-600 hover:bg-amber-700 text-white text-xs font-bold rounded-xl transition">Update Password</button>
                </div>
            </form>
        </div>
    </div>

    <!-- Edit User Modal -->
    <div id="user-edit-modal" class="hidden fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4">
        <div class="bg-white rounded-2xl shadow-2xl max-w-md w-full flex flex-col overflow-hidden">
            <div class="px-6 py-4 border-b border-slate-200 flex justify-between items-center bg-slate-50">
                <h3 class="font-bold text-slate-800 text-lg">Edit User Account</h3>
                <button type="button" onclick="closeUserEditModal()" class="text-slate-400 hover:text-slate-700 transition"><i class="ph ph-x text-xl"></i></button>
            </div>
            <div class="p-6">
                <form id="form-edit-user" onsubmit="handleSaveEditedUser(event)" class="space-y-4">
                    <input type="hidden" id="edit-user-id">
                    <div>
                        <label class="block text-sm font-semibold text-slate-700 mb-1">Username</label>
                        <input type="text" id="edit-user-name" required class="w-full px-4 py-2.5 border border-slate-300 rounded-xl text-sm outline-none focus:ring-2 focus:ring-blue-500 bg-white">
                    </div>
                    <div>
                        <label class="block text-sm font-semibold text-slate-700 mb-1">New Password <span class="text-slate-400 text-xs font-normal">(Leave blank to keep current)</span></label>
                        <input type="password" id="edit-user-pass" placeholder="••••••••" class="w-full px-4 py-2.5 border border-slate-300 rounded-xl text-sm outline-none focus:ring-2 focus:ring-blue-500 bg-white">
                    </div>
                    <div>
                        <label class="block text-sm font-semibold text-slate-700 mb-1">Role</label>
                        <select id="edit-user-role" required class="w-full px-4 py-2.5 border border-slate-300 rounded-xl text-sm outline-none focus:ring-2 focus:ring-blue-500 bg-white">
                            <option value="manager">Store Manager</option>
                            <option value="admin">Admin</option>
                        </select>
                    </div>
                    <div class="flex gap-3 pt-2">
                        <button type="button" onclick="closeUserEditModal()" class="flex-1 px-4 py-2.5 rounded-xl font-medium border border-slate-300 text-slate-700 hover:bg-slate-50 transition text-sm">Cancel</button>
                        <button type="submit" class="flex-1 px-4 py-2.5 rounded-xl font-semibold bg-blue-600 text-white hover:bg-blue-700 transition text-sm">Save Changes</button>
                    </div>
                </form>
            </div>
        </div>
    </div>

    <!-- Google Drive 2-Way Auto-Sync & Backup Modal -->
    <div id="drive-sync-modal" class="hidden fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4">
        <div class="bg-white rounded-2xl shadow-2xl max-w-lg w-full flex flex-col overflow-hidden">
            <div class="px-6 py-4 border-b border-slate-200 flex justify-between items-center bg-slate-50">
                <h3 class="font-bold text-slate-800 text-lg flex items-center gap-2"><i class="ph ph-cloud-arrow-up text-amber-600 text-xl"></i> Google Drive 2-Way Sync & Daily Backup</h3>
                <button type="button" onclick="closeDriveSyncModal()" class="text-slate-400 hover:text-slate-700 transition"><i class="ph ph-x text-xl"></i></button>
            </div>
            <div class="p-6 space-y-4 text-sm">
                <div class="p-3.5 bg-blue-50 text-blue-800 rounded-xl border border-blue-200 text-xs leading-relaxed">
                    <strong>How 2-Way Drive Sync Works:</strong><br>
                    1. Every change (new item, dispatch, or account update) is automatically pushed to your Google Drive folder.<br>
                    2. A dated Google Sheet / Excel backup is automatically created in your folder every day.<br>
                    3. You can pull/restore data from your Google Drive folder anytime across any computer.
                </div>

                <div>
                    <label class="block text-xs font-bold uppercase text-slate-600 mb-1">Google Drive Folder Link (Optional Reference)</label>
                    <input type="url" id="drive-folder-url" placeholder="https://drive.google.com/drive/folders/..." class="w-full px-3 py-2 border border-slate-300 rounded-xl text-xs outline-none focus:ring-2 focus:ring-blue-500">
                </div>

                <div>
                    <label class="block text-xs font-bold uppercase text-slate-600 mb-1">Google Apps Script Web App URL <span class="text-rose-500">*</span></label>
                    <input type="url" id="drive-webhook-url" placeholder="https://script.google.com/macros/s/.../exec" class="w-full px-3 py-2.5 border border-slate-300 rounded-xl text-sm outline-none focus:ring-2 focus:ring-blue-500">
                </div>

                <div class="flex items-center justify-between text-xs text-slate-500 pt-1 border-t border-slate-100">
                    <span>Last Synced: <strong id="last-drive-sync-display" class="text-slate-700">Never</strong></span>
                    <button type="button" onclick="openConnectedDriveFolder()" class="text-blue-600 hover:underline font-semibold flex items-center gap-1">
                        <i class="ph ph-folder-open"></i> Open Drive Folder
                    </button>
                </div>

                <div class="grid grid-cols-1 sm:grid-cols-3 gap-2 pt-2">
                    <button type="button" onclick="saveDriveWebhookUrl()" class="bg-slate-900 hover:bg-slate-800 text-white font-semibold py-2.5 px-3 rounded-xl transition text-xs">
                        Save Settings
                    </button>
                    <button type="button" onclick="syncToGoogleDrive(true)" class="bg-amber-600 hover:bg-amber-700 text-white font-semibold py-2.5 px-3 rounded-xl transition flex items-center justify-center gap-1 text-xs">
                        <i class="ph ph-cloud-arrow-up text-sm"></i> Push to Drive
                    </button>
                    <button type="button" onclick="loadFromGoogleDrive(true)" class="bg-blue-600 hover:bg-blue-700 text-white font-semibold py-2.5 px-3 rounded-xl transition flex items-center justify-center gap-1 text-xs">
                        <i class="ph ph-cloud-arrow-down text-sm"></i> Pull from Drive
                    </button>
                </div>
            </div>
        </div>
    </div>

    <!-- Admin Manage Models & Brands Modal -->
    <div id="manage-models-modal" class="hidden fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4">
        <div class="bg-white rounded-2xl shadow-2xl max-w-3xl w-full flex flex-col max-h-[90vh]">
            <div class="px-6 py-4 border-b border-slate-200 flex justify-between items-center bg-slate-50 rounded-t-2xl">
                <h3 class="font-bold text-slate-800 text-lg flex items-center gap-2"><i class="ph ph-gear-six text-indigo-600 text-xl"></i> Manage Brands & Equipment Models</h3>
                <button type="button" onclick="closeManageModelsModal()" class="text-slate-400 hover:text-slate-700 transition"><i class="ph ph-x text-xl"></i></button>
            </div>
            
            <div class="p-6 overflow-y-auto space-y-6">
                <!-- Add New Brand / Category Form -->
                <div class="bg-indigo-50/60 p-4 rounded-xl border border-indigo-200 space-y-3">
                    <h4 class="font-bold text-indigo-950 text-sm flex items-center gap-1.5"><i class="ph ph-folders text-indigo-600 text-base"></i> Add New Brand / Category Group</h4>
                    <form onsubmit="handleAddNewCategory(event)" class="flex gap-2">
                        <input type="text" id="new-category-name" required placeholder="e.g. Trimble Brand / Leica Systems..." class="flex-1 px-3 py-2 border border-slate-300 rounded-lg text-sm outline-none focus:border-indigo-500 bg-white">
                        <button type="submit" class="px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg text-sm font-semibold transition shrink-0">Add Brand Group</button>
                    </form>
                </div>

                <!-- Add New Model Form -->
                <div class="bg-slate-50 p-4 rounded-xl border border-slate-200 space-y-3">
                    <h4 class="font-bold text-slate-800 text-sm flex items-center gap-1.5"><i class="ph ph-plus-circle text-blue-600 text-base"></i> Add New Model to Existing Brand</h4>
                    <form id="form-add-model" onsubmit="handleAddNewModel(event)" class="grid grid-cols-1 sm:grid-cols-3 gap-3">
                        <div>
                            <label class="block text-xs font-semibold text-slate-600 mb-1">Select Brand / Category</label>
                            <select id="new-model-category" required class="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm bg-white outline-none focus:border-blue-500">
                                <!-- Populated dynamically -->
                            </select>
                        </div>
                        <div class="sm:col-span-2">
                            <label class="block text-xs font-semibold text-slate-600 mb-1">Model Name (e.g. JAVAD - TRIUMPH-4)</label>
                            <div class="flex gap-2">
                                <input type="text" id="new-model-name" required placeholder="Enter model name..." class="flex-1 px-3 py-2 border border-slate-300 rounded-lg text-sm outline-none focus:border-blue-500 bg-white">
                                <button type="submit" class="px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-sm font-semibold transition shrink-0">Add Model</button>
                            </div>
                        </div>
                    </form>
                </div>

                <!-- Existing Models List -->
                <div>
                    <h4 class="font-bold text-slate-800 text-sm mb-3">Existing Models (Edit name inline and click 'Save')</h4>
                    <div id="models-management-list" class="space-y-4">
                        <!-- Populated dynamically -->
                    </div>
                </div>
            </div>

            <div class="px-6 py-3 bg-slate-50 border-t border-slate-200 flex justify-end">
                <button type="button" onclick="closeManageModelsModal()" class="px-5 py-2 bg-slate-900 text-white rounded-xl text-sm font-semibold hover:bg-black transition">Done</button>
            </div>
        </div>
    </div>

    <!-- Edit Item Modal -->
    <div id="edit-modal" class="hidden fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4">
        <div class="bg-white rounded-2xl shadow-2xl max-w-md w-full flex flex-col">
            <div class="px-6 py-4 border-b border-slate-200 flex justify-between items-center">
                <h3 class="font-bold text-slate-800 text-lg">Edit Equipment</h3>
                <button type="button" onclick="closeEditModal()" class="text-slate-400 hover:text-slate-700 transition"><i class="ph ph-x text-xl"></i></button>
            </div>
            <div class="p-6">
                <form id="form-edit" class="space-y-4">
                    <input type="hidden" id="edit-id">
                    <div>
                        <label class="block text-sm font-semibold text-slate-700 mb-1">Equipment Name / Category</label>
                        <select id="edit-name" required class="w-full px-4 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none bg-white">
                            <!-- Populated dynamically -->
                        </select>
                    </div>
                    <div>
                        <label class="block text-sm font-semibold text-slate-700 mb-1">Received From (Source)</label>
                        <input type="text" id="edit-source" required class="w-full px-4 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none bg-white">
                    </div>
                    <div>
                        <label class="block text-sm font-semibold text-slate-700 mb-1">Serial Number</label>
                        <input type="text" id="edit-sn" required class="w-full px-4 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none uppercase">
                    </div>
                    <div>
                        <label class="block text-sm font-semibold text-slate-700 mb-1">Accessories Included</label>
                        <input type="text" id="edit-accessories" class="w-full px-4 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none">
                    </div>
                    <div class="flex gap-3 pt-2">
                        <button type="button" onclick="closeEditModal()" class="flex-1 px-4 py-2.5 rounded-lg font-medium border border-slate-300 text-slate-700 hover:bg-slate-50 transition">Cancel</button>
                        <button type="submit" class="flex-1 px-4 py-2.5 rounded-lg font-medium bg-blue-600 text-white hover:bg-blue-700 transition">Save Changes</button>
                    </div>
                </form>
            </div>
        </div>
    </div>

    <!-- CSV Import Modal -->
    <div id="csv-import-modal" class="hidden fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4">
        <div class="bg-white rounded-2xl shadow-2xl max-w-5xl w-full flex flex-col max-h-[90vh]">
            <div class="px-6 py-4 border-b border-slate-200 flex justify-between items-center bg-slate-50 rounded-t-2xl">
                <h3 class="font-bold text-slate-800 flex items-center gap-2"><i class="ph ph-file-arrow-up text-xl text-blue-600"></i> Import Horizontal CSV Format</h3>
                <button type="button" onclick="closeCsvModal()" class="text-slate-400 hover:text-slate-700 transition"><i class="ph ph-x text-xl"></i></button>
            </div>
            <div class="p-6 overflow-y-auto space-y-6">
                <div>
                    <label class="block text-sm font-semibold text-slate-700 mb-2">Upload CSV File</label>
                    <input type="file" id="csv-file-input" accept=".csv" class="w-full text-sm text-slate-500 file:mr-4 file:py-2.5 file:px-4 file:rounded-xl file:border-0 file:text-sm file:font-semibold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100 border border-slate-300 rounded-xl p-2">
                    <p class="text-xs text-slate-500 mt-1">Expected format: Model names in row 1, with serial numbers listed vertically below each column.</p>
                </div>

                <div class="flex flex-wrap gap-3">
                    <button type="button" onclick="downloadCsvTemplate()" class="px-4 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-sm font-medium transition flex items-center gap-2">
                        <i class="ph ph-download"></i> Download Horizontal Sample Template
                    </button>
                    <button type="button" onclick="processCsvUpload()" class="px-5 py-2.5 bg-blue-600 hover:bg-blue-700 text-white rounded-xl text-sm font-semibold transition flex items-center gap-2">
                        <i class="ph ph-upload-simple"></i> Process & Import Data
                    </button>
                </div>
            </div>
        </div>
    </div>

    <!-- Track Journey Modal -->
    <div id="track-modal" class="hidden fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4">
        <div class="bg-white rounded-2xl shadow-2xl max-w-lg w-full flex flex-col overflow-hidden">
            <div class="px-6 py-4 border-b border-slate-200 flex justify-between items-center bg-slate-50">
                <h3 class="font-bold text-slate-800 text-lg flex items-center gap-2"><i class="ph ph-compass text-blue-600 text-xl"></i> Material Journey Timeline</h3>
                <button type="button" onclick="closeTrackModal()" class="text-slate-400 hover:text-slate-700 transition"><i class="ph ph-x text-xl"></i></button>
            </div>
            <div class="p-6 max-h-[75vh] overflow-y-auto" id="track-modal-content">
                <!-- Timeline content -->
            </div>
            <div class="px-6 py-3 bg-slate-50 border-t border-slate-200 flex justify-end">
                <button type="button" onclick="closeTrackModal()" class="px-4 py-2 bg-slate-900 text-white rounded-xl text-sm font-semibold hover:bg-black transition">Close</button>
            </div>
        </div>
    </div>

    <!-- Dispatch Full Details Modal -->
    <div id="dispatch-details-modal" class="hidden fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4">
        <div class="bg-white rounded-2xl shadow-2xl max-w-lg w-full flex flex-col overflow-hidden">
            <div class="px-6 py-4 border-b border-slate-200 flex justify-between items-center bg-slate-50">
                <h3 class="font-bold text-slate-800 text-lg flex items-center gap-2"><i class="ph ph-info text-blue-600 text-xl"></i> Dispatch Audit & Approval Details</h3>
                <button type="button" onclick="closeDispatchDetailsModal()" class="text-slate-400 hover:text-slate-700 transition"><i class="ph ph-x text-xl"></i></button>
            </div>
            <div class="p-6 max-h-[75vh] overflow-y-auto space-y-4 text-sm" id="dispatch-details-content">
                <!-- Populated by JS -->
            </div>
            <div class="px-6 py-3 bg-slate-50 border-t border-slate-200 flex justify-end">
                <button type="button" onclick="closeDispatchDetailsModal()" class="px-4 py-2 bg-slate-900 text-white rounded-xl text-sm font-semibold hover:bg-black transition">Close</button>
            </div>
        </div>
    </div>

    <!-- Approval Notes & Proof Modal -->
    <div id="approval-details-modal" class="hidden fixed inset-0 z-50 bg-slate-900/70 backdrop-blur-sm flex items-center justify-center p-4">
        <div class="bg-white rounded-2xl shadow-2xl max-w-md w-full flex flex-col overflow-hidden">
            <div class="px-6 py-4 border-b border-slate-200 flex justify-between items-center bg-slate-50">
                <h3 class="font-bold text-slate-800 text-lg">Approval Notes & Verification Proof</h3>
                <button type="button" onclick="closeApprovalDetailsModal()" class="text-slate-400 hover:text-slate-700 transition"><i class="ph ph-x text-xl"></i></button>
            </div>
            <div class="p-6 space-y-4">
                <input type="hidden" id="approval-item-id">
                <div>
                    <label class="block text-sm font-semibold text-slate-700 mb-1">Admin Approval Notes / Proof Description <span class="text-rose-500">*</span></label>
                    <textarea id="approval-notes" rows="3" required placeholder="e.g. Verified client purchase order..." 
                        class="w-full px-4 py-3 border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none transition text-sm"></textarea>
                </div>
                <button type="button" onclick="proceedToCameraCapture()" class="w-full bg-blue-600 text-white font-bold py-3 px-4 rounded-xl hover:bg-blue-700 transition flex justify-center items-center gap-2">
                    <i class="ph-fill ph-camera text-xl"></i> Next: Take Verification Photo
                </button>
            </div>
        </div>
    </div>

    <!-- Camera Modal -->
    <div id="camera-modal" class="hidden fixed inset-0 z-50 bg-slate-900/90 backdrop-blur-sm flex items-center justify-center p-4">
        <div class="bg-white rounded-2xl shadow-2xl max-w-md w-full flex flex-col overflow-hidden">
            <div class="px-6 py-4 border-b border-slate-200 flex justify-between items-center bg-white">
                <h3 class="font-bold text-slate-800 text-lg">Capture Proof Photo</h3>
                <button type="button" onclick="closeCameraModal()" class="text-slate-400 hover:text-slate-700 transition"><i class="ph ph-x text-xl"></i></button>
            </div>
            <div class="p-4 bg-black flex justify-center">
                <video id="camera-stream" autoplay playsinline class="w-full max-w-sm rounded-lg object-cover h-64"></video>
                <canvas id="camera-canvas" class="hidden"></canvas>
            </div>
            <div class="p-6 bg-white flex flex-col gap-2 justify-center">
                <button type="button" onclick="capturePhoto()" class="bg-emerald-600 text-white font-bold py-3 px-8 rounded-full shadow-lg hover:bg-emerald-700 transition flex items-center justify-center gap-2">
                    <i class="ph-fill ph-check-circle text-xl"></i> Snap & Finalize Approval
                </button>
                <button type="button" onclick="skipPhotoAndApprove()" class="text-xs text-slate-500 hover:text-slate-800 underline py-1 text-center">
                    Skip Photo & Approve Directly
                </button>
            </div>
        </div>
    </div>

    <!-- Gatepass Modal Preview -->
    <div id="gatepass-preview-modal" class="hidden fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4">
        <div class="bg-white rounded-2xl shadow-2xl max-w-3xl w-full flex flex-col max-h-[90vh]">
            <div class="px-6 py-4 border-b border-slate-200 flex justify-between items-center bg-slate-50 rounded-t-2xl">
                <h3 class="font-bold text-slate-800 flex items-center gap-2"><i class="ph ph-file-text text-xl"></i> Gatepass Preview</h3>
                <button type="button" onclick="closeGatepassPreview()" class="text-slate-400 hover:text-slate-700 transition"><i class="ph ph-x text-xl"></i></button>
            </div>
            
            <div class="p-4 md:p-8 overflow-y-auto bg-slate-100 flex-1 flex justify-center">
                <div class="bg-white shadow-md p-6 md:p-8 w-full max-w-[21cm] min-h-[14cm] border border-slate-300 relative text-black text-sm md:text-base" id="preview-document">
                    <!-- Cloned via JS -->
                </div>
            </div>

            <div class="px-6 py-4 border-t border-slate-200 bg-white rounded-b-2xl flex justify-end gap-3">
                <button type="button" onclick="closeGatepassPreview()" class="px-5 py-2.5 rounded-xl font-medium text-slate-600 hover:bg-slate-100 transition">Done</button>
                <button type="button" onclick="window.print()" class="px-5 py-2.5 rounded-xl font-medium bg-slate-900 text-white hover:bg-black transition flex items-center gap-2">
                    <i class="ph ph-printer"></i> Print Document
                </button>
            </div>
        </div>
    </div>

    <!-- Printable Gatepass View -->
    <div id="gatepass-print-view" class="bg-white text-black font-sans">
        <div class="text-center mb-6 border-b-2 border-black pb-4">
            <div class="flex justify-center mb-2">
                <h2 class="text-2xl font-black text-red-600 uppercase tracking-widest">JANAK POSITIONING</h2>
            </div>
            <p class="text-xs mt-1">Corporate Office: Sector 62, Noida, UP, India</p>
            <h2 class="text-xl font-bold mt-3 underline">DELIVERY CHALLAN / GATEPASS</h2>
        </div>

        <div class="flex justify-between mb-6 text-sm md:text-base border-b border-black pb-4">
            <div>
                <p><span class="font-bold">Gatepass No:</span> <span id="print-gp-no"></span></p>
                <p class="mt-1"><span class="font-bold">Date:</span> <span id="print-date"></span></p>
                <p class="mt-1"><span class="font-bold">Challan/Invoice No:</span> <span id="print-challan"></span></p>
            </div>
            <div class="text-right">
                <p class="font-bold text-base" id="print-company"></p>
                <p id="print-to" class="mt-1 text-sm"></p>
                <p class="mt-1 text-sm"><span class="font-bold">Phone:</span> <span id="print-phone"></span></p>
                <p id="print-courier-info" class="mt-1 text-sm font-semibold text-blue-800 hidden"></p>
            </div>
        </div>

        <table class="w-full mb-6 border-collapse border border-black text-sm md:text-base">
            <thead>
                <tr class="bg-gray-100">
                    <th class="border border-black px-4 py-2 text-center w-12">S.No</th>
                    <th class="border border-black px-4 py-2 text-left">Equipment Model / Category</th>
                    <th class="border border-black px-4 py-2 text-left w-1/3">Serial Number</th>
                    <th class="border border-black px-4 py-2 text-center w-24">Qty</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td class="border border-black px-4 py-3 text-center">1</td>
                    <td class="border border-black px-4 py-3 font-semibold" id="print-item-name"></td>
                    <td class="border border-black px-4 py-3 font-mono" id="print-item-sn"></td>
                    <td class="border border-black px-4 py-3 text-center">1</td>
                </tr>
            </tbody>
        </table>

        <div class="mb-12 text-sm md:text-base space-y-3">
            <div>
                <p class="font-bold mb-1">Accessories / Full Details:</p>
                <div id="print-remarks" class="p-3 border border-black min-h-[60px] italic bg-gray-50"></div>
            </div>
        </div>

        <div class="flex justify-between mt-24 px-4 md:px-12 text-sm md:text-base">
            <div class="text-center">
                <div class="border-t-2 border-black w-40 md:w-48 pt-2">Authorized Signatory</div>
            </div>
            <div class="text-center">
                <div class="border-t-2 border-black w-40 md:w-48 pt-2">Receiver's Signature</div>
            </div>
        </div>
    </div>

    <!-- Floating Toast Container -->
    <div id="toast-container" class="fixed bottom-6 right-6 z-50 flex flex-col gap-2"></div>

    <!-- JavaScript Application Logic -->
    <script>
        const STORAGE_KEY = 'janak_stock_inventory';
        const CATALOG_KEY = 'janak_catalog_models';
        const AUTH_KEY = 'janak_auth_role';
        const USERS_KEY = 'janak_system_users';
        const LOGGED_USER_KEY = 'janak_logged_in_username';
        const DRIVE_WEBHOOK_KEY = 'janak_drive_webhook_url';
        const DRIVE_FOLDER_LINK_KEY = 'janak_drive_folder_url';
        const LAST_SYNC_DATE_KEY = 'janak_last_drive_sync_date';
        const ADMIN_EMAIL = 'anant@janaksurvey.in';
        const MIN_STOCK_THRESHOLD = 5;

        const DEFAULT_CATALOG = {
            "JAVAD - Smart Antennas": [
                "JAVAD - TRIUMPH-3NR", "JAVAD - TRIUMPH-1M", "JAVAD - TRIUMPH-1M Plus", 
                "JAVAD - TRIUMPH-2", "JAVAD - TRIUMPH-3", "JAVAD - TRIUMPH-LS Plus", "JAVAD - PulserOne"
            ],
            "JAVAD - OEM GNSS Boards": ["JAVAD - TR-2S OEM Board", "JAVAD - TRE-3S OEM Board", "JAVAD - TRH Heading OEM Board"],
            "JAVAD - GNSS Antennas": ["JAVAD - GrAnt-G3T-JS", "JAVAD - GrAnt-G3-JS", "JAVAD - GrAnt-G2T-JSM", "JAVAD - GrAnt-3L", "JAVAD - GrAnt-3LS", "JAVAD - SpaceAnt", "JAVAD - MiniAnt-G3", "JAVAD - MiniAnt-G3T"],
            "JAVAD - INS Solutions": ["JAVAD - ALPHA INS", "JAVAD - TR-2S INS", "JAVAD - TR-2S AJ INS", "JAVAD - TR-2S HDA INS", "JAVAD - TR-3S DUO INS"],
            "FOIF Brand": ["FOIF - A90", "FOIF - A66", "FOIF - RTS 102", "FOIF - RTS 392", "FOIF - RTS350"],
            "Alpha Geo Brand": ["Alpha Geo - GNSS Receiver", "Alpha Geo - Slam"],
            "Janak Brand": ["Janak Brand - GNSS Receiver", "Janak Brand - Controller"]
        };

        const DEFAULT_USERS = [
            { id: 'usr-1', username: 'admin', password: 'admin123', role: 'admin', created: '2026-01-01' },
            { id: 'usr-2', username: 'manager', password: 'manager123', role: 'manager', created: '2026-01-01' }
        ];

        let inventory = [];
        let catalog = {};
        let usersList = [];
        let currentUserRole = null;
        let currentUsername = null;
        let streamRef = null;

        function checkAuth() {
            const savedRole = localStorage.getItem(AUTH_KEY);
            const savedUser = localStorage.getItem(LOGGED_USER_KEY);
            if ((savedRole === 'admin' || savedRole === 'manager') && savedUser) {
                currentUserRole = savedRole;
                currentUsername = savedUser;
                showApp();
            } else {
                document.getElementById('login-screen').classList.remove('hidden');
                document.getElementById('app-container').classList.add('hidden');
                document.getElementById('app-container').classList.remove('flex');
            }
        }

        function selectLoginRole(role) {
            const adminBtn = document.getElementById('role-btn-admin');
            const managerBtn = document.getElementById('role-btn-manager');
            const adminIcon = document.getElementById('role-icon-admin');
            const managerIcon = document.getElementById('role-icon-manager');

            if (role === 'admin') {
                adminBtn.className = "py-2.5 px-3 bg-slate-900 text-white border border-slate-900 text-xs font-bold rounded-xl transition flex items-center justify-center gap-1.5 shadow-md scale-105";
                adminIcon.className = "ph-fill ph-shield-check text-emerald-400 text-base";
                managerBtn.className = "py-2.5 px-3 bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-200 text-xs font-bold rounded-xl transition flex items-center justify-center gap-1.5 shadow-sm";
                managerIcon.className = "ph-fill ph-user-circle text-slate-500 text-base";
                document.getElementById('login-user').value = 'admin';
                document.getElementById('login-pass').value = 'admin123';
            } else {
                managerBtn.className = "py-2.5 px-3 bg-blue-600 text-white border border-blue-600 text-xs font-bold rounded-xl transition flex items-center justify-center gap-1.5 shadow-md scale-105";
                managerIcon.className = "ph-fill ph-user-circle text-white text-base";
                adminBtn.className = "py-2.5 px-3 bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-200 text-xs font-bold rounded-xl transition flex items-center justify-center gap-1.5 shadow-sm";
                adminIcon.className = "ph-fill ph-shield-check text-slate-500 text-base";
                document.getElementById('login-user').value = 'manager';
                document.getElementById('login-pass').value = 'manager123';
            }
        }

        function handleLogin(e) {
            e.preventDefault();
            const inputUser = document.getElementById('login-user').value.trim();
            const inputPass = document.getElementById('login-pass').value.trim();
            loadUsers();

            const foundUser = usersList.find(u => u.username.toLowerCase() === inputUser.toLowerCase() && u.password === inputPass);

            if (foundUser) {
                currentUserRole = foundUser.role;
                currentUsername = foundUser.username;
            } else if (inputUser.toLowerCase() === 'admin' && inputPass === 'admin123') {
                currentUserRole = 'admin';
                currentUsername = 'admin';
            } else if (inputUser.toLowerCase() === 'manager' && inputPass === 'manager123') {
                currentUserRole = 'manager';
                currentUsername = 'manager';
            } else {
                document.getElementById('login-error').classList.remove('hidden');
                return;
            }

            localStorage.setItem(AUTH_KEY, currentUserRole);
            localStorage.setItem(LOGGED_USER_KEY, currentUsername);
            document.getElementById('login-error').classList.add('hidden');
            document.getElementById('form-login').reset();
            showApp();
        }

        function logout() {
            currentUserRole = null;
            currentUsername = null;
            localStorage.removeItem(AUTH_KEY);
            localStorage.removeItem(LOGGED_USER_KEY);
            document.getElementById('app-container').classList.add('hidden');
            document.getElementById('app-container').classList.remove('flex');
            document.getElementById('login-screen').classList.remove('hidden');
            showToast("Logged out successfully");
        }

        function openForgotPasswordModal() {
            document.getElementById('recovery-pin').value = '';
            document.getElementById('recovery-new-pass').value = '';
            document.getElementById('recovery-error').classList.add('hidden');
            document.getElementById('forgot-password-modal').classList.remove('hidden');
            document.getElementById('recovery-pin').focus();
        }

        function closeForgotPasswordModal() {
            document.getElementById('forgot-password-modal').classList.add('hidden');
        }

        function handlePasswordRecovery(e) {
            e.preventDefault();
            const pin = document.getElementById('recovery-pin').value.trim();
            const newUser = document.getElementById('recovery-new-user').value.trim();
            const newPass = document.getElementById('recovery-new-pass').value.trim();

            if (pin === '8317') {
                loadUsers();
                let adminIndex = usersList.findIndex(u => u.role === 'admin');
                if (adminIndex > -1) {
                    usersList[adminIndex].username = newUser;
                    usersList[adminIndex].password = newPass;
                } else {
                    usersList.push({ id: 'USR-ADMIN', username: newUser, password: newPass, role: 'admin', created: new Date().toISOString().split('T')[0] });
                }
                saveUsers();
                closeForgotPasswordModal();
                showToast("Admin password successfully reset!");
                document.getElementById('login-user').value = newUser;
                document.getElementById('login-pass').value = newPass;
                selectLoginRole('admin');
            } else {
                document.getElementById('recovery-error').classList.remove('hidden');
                document.getElementById('recovery-pin').value = '';
                document.getElementById('recovery-pin').focus();
            }
        }

        function showApp() {
            document.getElementById('login-screen').classList.add('hidden');
            const appContainer = document.getElementById('app-container');
            appContainer.classList.remove('hidden');
            appContainer.classList.add('flex');
            document.getElementById('user-role-display').textContent = `${currentUsername} (${currentUserRole === 'admin' ? 'Admin' : 'Manager'})`;
            
            const manageModelsBtn = document.getElementById('admin-manage-models-btn');
            const navManageModels = document.getElementById('nav-manage-models');
            const navUserManagement = document.getElementById('nav-user-management');
            const inTabManageBtn = document.getElementById('in-tab-manage-btn');

            if (currentUserRole === 'admin') {
                document.getElementById('nav-approvals').classList.remove('hidden');
                navUserManagement.classList.remove('hidden');
                navManageModels.classList.remove('hidden');
                manageModelsBtn.classList.remove('hidden');
                manageModelsBtn.classList.add('flex');
                inTabManageBtn.classList.remove('hidden');
                inTabManageBtn.classList.add('flex');
            } else {
                document.getElementById('nav-approvals').classList.add('hidden');
                navUserManagement.classList.add('hidden');
                navManageModels.classList.add('hidden');
                manageModelsBtn.classList.add('hidden');
                manageModelsBtn.classList.remove('flex');
                inTabManageBtn.classList.add('hidden');
                inTabManageBtn.classList.remove('flex');
            }
            initApp();
        }

        async function initApp() {
            loadCatalog();
            loadUsers();
            loadData();
            document.getElementById('in-date').valueAsDate = new Date();
            populateModelDropdowns();
            renderTable();
            renderUsersTable();
            updateMaterialOutSelect();
            updateStats();
            checkLowStockLevels();
            updateDriveSyncBadge();

            const webhookUrl = localStorage.getItem(DRIVE_WEBHOOK_KEY);
            if (webhookUrl) {
                await loadFromGoogleDrive(false);
                checkAutoDailyDriveBackup();
            }
        }

        function loadCatalog() {
            const data = localStorage.getItem(CATALOG_KEY);
            catalog = data ? JSON.parse(data) : JSON.parse(JSON.stringify(DEFAULT_CATALOG));
        }

        function saveCatalog() {
            localStorage.setItem(CATALOG_KEY, JSON.stringify(catalog));
            syncToGoogleDrive(false);
        }

        function loadUsers() {
            const data = localStorage.getItem(USERS_KEY);
            usersList = data ? JSON.parse(data) : JSON.parse(JSON.stringify(DEFAULT_USERS));
        }

        function saveUsers() {
            localStorage.setItem(USERS_KEY, JSON.stringify(usersList));
            syncToGoogleDrive(false);
        }

        function loadData() {
            const data = localStorage.getItem(STORAGE_KEY);
            inventory = data ? JSON.parse(data) : [];
        }

        function saveData() {
            localStorage.setItem(STORAGE_KEY, JSON.stringify(inventory));
            syncToGoogleDrive(false);
        }

        function requestManageModelsAccess() {
            document.getElementById('model-pin-input').value = '';
            document.getElementById('pin-error').classList.add('hidden');
            document.getElementById('model-pin-modal').classList.remove('hidden');
            document.getElementById('model-pin-input').focus();
        }

        function closeModelPinModal() {
            document.getElementById('model-pin-modal').classList.add('hidden');
        }

        function verifyModelPin(e) {
            e.preventDefault();
            if (document.getElementById('model-pin-input').value.trim() === '8317') {
                closeModelPinModal();
                openManageModelsModal();
            } else {
                document.getElementById('pin-error').classList.remove('hidden');
            }
        }

        function updateDriveSyncBadge(state = 'idle') {
            const webhookUrl = localStorage.getItem(DRIVE_WEBHOOK_KEY);
            const badge = document.getElementById('drive-status-badge');
            const dot = document.getElementById('drive-status-dot');
            const text = document.getElementById('drive-status-text');

            if (!webhookUrl) {
                badge.className = "px-2.5 py-1 rounded-full text-[11px] font-bold border flex items-center gap-1.5 transition bg-slate-100 text-slate-600 border-slate-200";
                dot.className = "w-2 h-2 rounded-full bg-slate-400";
                text.textContent = "Drive: Setup Needed";
                return;
            }
            if (state === 'syncing') {
                badge.className = "px-2.5 py-1 rounded-full text-[11px] font-bold border flex items-center gap-1.5 transition bg-blue-50 text-blue-700 border-blue-200";
                dot.className = "w-2 h-2 rounded-full bg-blue-500 animate-ping";
                text.textContent = "Syncing with Drive...";
            } else {
                badge.className = "px-2.5 py-1 rounded-full text-[11px] font-bold border flex items-center gap-1.5 transition bg-emerald-50 text-emerald-700 border-emerald-200";
                dot.className = "w-2 h-2 rounded-full bg-emerald-500";
                text.textContent = "Drive Synced";
            }
        }

        function openDriveSyncModal() {
            document.getElementById('drive-webhook-url').value = localStorage.getItem(DRIVE_WEBHOOK_KEY) || '';
            document.getElementById('drive-folder-url').value = localStorage.getItem(DRIVE_FOLDER_LINK_KEY) || '';
            document.getElementById('last-drive-sync-display').textContent = localStorage.getItem(LAST_SYNC_DATE_KEY) || 'Never';
            document.getElementById('drive-sync-modal').classList.remove('hidden');
        }

        function closeDriveSyncModal() {
            document.getElementById('drive-sync-modal').classList.add('hidden');
        }

        function openConnectedDriveFolder() {
            const folderLink = document.getElementById('drive-folder-url').value.trim() || localStorage.getItem(DRIVE_FOLDER_LINK_KEY);
            if (folderLink) window.open(folderLink, '_blank');
            else showToast("Please paste your Google Drive folder link first.", "info");
        }

        function saveDriveWebhookUrl() {
            localStorage.setItem(DRIVE_WEBHOOK_KEY, document.getElementById('drive-webhook-url').value.trim());
            localStorage.setItem(DRIVE_FOLDER_LINK_KEY, document.getElementById('drive-folder-url').value.trim());
            updateDriveSyncBadge();
            showToast("Google Drive Sync settings saved!");
        }

        function checkAutoDailyDriveBackup() {
            const webhookUrl = localStorage.getItem(DRIVE_WEBHOOK_KEY);
            if (!webhookUrl || inventory.length === 0) return;
            const todayStr = new Date().toISOString().split('T')[0];
            const lastSync = localStorage.getItem(LAST_SYNC_DATE_KEY);
            if (!lastSync || !lastSync.startsWith(todayStr)) syncToGoogleDrive(false);
        }

        async function syncToGoogleDrive(manual = false) {
            const webhookUrl = localStorage.getItem(DRIVE_WEBHOOK_KEY);
            if (!webhookUrl) { if (manual) showToast("Please configure Web App URL first.", "error"); return; }
            updateDriveSyncBadge('syncing');
            try {
                await fetch(webhookUrl, {
                    method: 'POST', mode: 'no-cors',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ catalog, users: usersList, inventory, timestamp: new Date().toISOString() })
                });
                const nowStr = new Date().toISOString().split('T')[0] + ' ' + new Date().toLocaleTimeString();
                localStorage.setItem(LAST_SYNC_DATE_KEY, nowStr);
                document.getElementById('last-drive-sync-display').textContent = nowStr;
                updateDriveSyncBadge('idle');
                if (manual) showToast("Synchronized with Google Drive successfully!");
            } catch (err) {
                updateDriveSyncBadge('error');
                if (manual) showToast("Failed to reach Google Drive webhook.", "error");
            }
        }

        async function loadFromGoogleDrive(manual = false) {
            const webhookUrl = localStorage.getItem(DRIVE_WEBHOOK_KEY);
            if (!webhookUrl) return;
            try {
                const response = await fetch(webhookUrl, { method: 'GET' });
                const data = await response.json();
                if (data && data.status === 'success') {
                    if (data.catalog) catalog = data.catalog;
                    if (data.users) usersList = data.users;
                    if (data.inventory) inventory = data.inventory;
                    saveData(); saveCatalog(); saveUsers();
                    populateModelDropdowns(); renderTable(); updateStats();
                }
            } catch (err) {}
        }

        function exportToExcel() {
            if (inventory.length === 0) { showToast("No inventory data to export", "error"); return; }
            const masterData = inventory.map(item => ({
                "Category": getCategoryFromName(item.name),
                "Equipment Model": item.name || "",
                "Serial Number": item.sn || "",
                "Status": item.status || "",
                "Received From": item.source || "",
                "Date Received": item.dateIn || "",
                "Client": item.company || "",
                "Challan": item.challan || "",
                "Date Dispatched": item.dateOut ? formatDate(item.dateOut) : ""
            }));
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(masterData), "Master_Inventory");
            XLSX.writeFile(wb, `Janak_Stock_Backup_${new Date().toISOString().split('T')[0]}.xlsx`);
            showToast("Excel backup downloaded!");
        }

        function populateModelDropdowns() {
            [document.getElementById('in-name'), document.getElementById('edit-name')].forEach(select => {
                if (!select) return;
                const currentVal = select.value;
                select.innerHTML = '<option value="" disabled selected>Select equipment model...</option>';
                for (const [category, models] of Object.entries(catalog)) {
                    const optgroup = document.createElement('optgroup');
                    optgroup.label = category;
                    models.forEach(model => {
                        const option = document.createElement('option');
                        option.value = model; option.textContent = model;
                        optgroup.appendChild(option);
                    });
                    select.appendChild(optgroup);
                }
                if (currentVal) select.value = currentVal;
            });
            const catSelect = document.getElementById('new-model-category');
            if (catSelect) {
                catSelect.innerHTML = '';
                Object.keys(catalog).forEach(cat => {
                    const opt = document.createElement('option');
                    opt.value = cat; opt.textContent = cat;
                    catSelect.appendChild(opt);
                });
            }
        }

        function openManageModelsModal() {
            populateModelDropdowns();
            renderManageModelsList();
            document.getElementById('manage-models-modal').classList.remove('hidden');
        }

        function closeManageModelsModal() {
            document.getElementById('manage-models-modal').classList.add('hidden');
            populateModelDropdowns(); renderTable(); updateStats(); updateMaterialOutSelect();
        }

        function renderManageModelsList() {
            const container = document.getElementById('models-management-list');
            container.innerHTML = '';
            Object.entries(catalog).forEach(([category, models], catIdx) => {
                let modelsHtml = '';
                models.forEach((model, modIdx) => {
                    const inputId = `model-input-${catIdx}-${modIdx}`;
                    modelsHtml += `
                        <div class="flex items-center justify-between gap-2 bg-white p-2 rounded-lg border text-xs mb-1">
                            <input type="text" id="${inputId}" value="${model}" class="flex-1 font-mono px-2 py-1 border rounded outline-none">
                            <button type="button" onclick="saveInlineModelRename(${catIdx}, ${modIdx}, '${inputId}')" class="px-3 py-1 bg-blue-600 text-white rounded font-semibold">Save</button>
                            <button type="button" onclick="deleteCatalogModel(${catIdx}, ${modIdx})" class="px-2 py-1 bg-rose-50 text-rose-600 rounded">Delete</button>
                        </div>`;
                });
                const div = document.createElement('div');
                div.className = 'bg-slate-50 p-3 rounded-xl border';
                div.innerHTML = `<h5 class="font-bold text-xs text-blue-700 uppercase mb-2">${category}</h5><div class="space-y-1">${modelsHtml}</div>`;
                container.appendChild(div);
            });
        }

        function handleAddNewCategory(e) {
            e.preventDefault();
            const newCat = document.getElementById('new-category-name').value.trim();
            if (newCat && !catalog[newCat]) {
                catalog[newCat] = [];
                saveCatalog();
                document.getElementById('new-category-name').value = '';
                renderManageModelsList(); populateModelDropdowns();
                showToast("Brand group added!");
            }
        }

        function handleAddNewModel(e) {
            e.preventDefault();
            const category = document.getElementById('new-model-category').value;
            const modelName = document.getElementById('new-model-name').value.trim();
            if (modelName && catalog[category] && !catalog[category].includes(modelName)) {
                catalog[category].push(modelName);
                saveCatalog();
                document.getElementById('new-model-name').value = '';
                renderManageModelsList(); populateModelDropdowns();
                showToast("Model added!");
            }
        }

        function saveInlineModelRename(catIdx, modIdx, inputId) {
            const categories = Object.keys(catalog);
            const category = categories[catIdx];
            const oldName = catalog[category][modIdx];
            const newName = document.getElementById(inputId).value.trim();
            if (newName && newName !== oldName) {
                catalog[category][modIdx] = newName;
                inventory.forEach(i => { if (i.name === oldName) i.name = newName; });
                saveCatalog(); saveData();
                renderManageModelsList(); renderTable(); updateStats();
                showToast("Model renamed successfully!");
            }
        }

        function deleteCatalogModel(catIdx, modIdx) {
            const categories = Object.keys(catalog);
            catalog[categories[catIdx]].splice(modIdx, 1);
            saveCatalog(); renderManageModelsList(); populateModelDropdowns();
            showToast("Model deleted.");
        }

        function checkLowStockLevels() {
            const modelCounts = {};
            inventory.forEach(i => { if (i.status === 'In') modelCounts[i.name] = (modelCounts[i.name] || 0) + 1; });
            const widgetContainer = document.getElementById('low-stock-widget-container');
            const lowStockList = document.getElementById('low-stock-list');
            lowStockList.innerHTML = '';
            let lowStockModels = [];

            Object.entries(modelCounts).forEach(([model, count]) => {
                if (count <= MIN_STOCK_THRESHOLD) {
                    lowStockModels.push({ model, count });
                }
            });

            if (lowStockModels.length > 0) {
                widgetContainer.classList.remove('hidden');
                lowStockModels.forEach(item => {
                    const div = document.createElement('div');
                    div.className = 'bg-white p-3 rounded-xl border border-rose-200 shadow-sm flex justify-between items-center';
                    div.innerHTML = `<span class="font-semibold text-slate-800 text-xs">${item.model}</span><span class="px-2 py-0.5 bg-rose-100 text-rose-700 rounded text-xs font-bold">${item.count} left</span>`;
                    lowStockList.appendChild(div);
                });
            } else {
                widgetContainer.classList.add('hidden');
            }
        }

        function showToast(message, type = 'success') {
            const container = document.getElementById('toast-container');
            const toast = document.createElement('div');
            toast.className = `flex items-center gap-3 px-5 py-3 rounded-xl shadow-lg border text-sm font-medium transform transition-all duration-300`;
            if (type === 'success') {
                toast.classList.add('bg-slate-900', 'text-white', 'border-slate-800');
                toast.innerHTML = `<i class="ph-fill ph-check-circle text-emerald-400 text-xl"></i> ${message}`;
            } else {
                toast.classList.add('bg-white', 'text-rose-600', 'border-rose-200');
                toast.innerHTML = `<i class="ph-fill ph-warning-circle text-rose-500 text-xl"></i> ${message}`;
            }
            container.appendChild(toast);
            setTimeout(() => toast.remove(), 3000);
        }

        function switchTab(tabId) {
            if (tabId === 'user-management' && currentUserRole !== 'admin') {
                showToast("Access restricted to Administrators only.", "error");
                return;
            }
            document.querySelectorAll('.nav-btn').forEach(btn => {
                btn.classList.remove('bg-blue-600', 'text-white');
                btn.classList.add('text-slate-400', 'hover:bg-slate-800');
            });
            const activeBtn = document.getElementById(`nav-${tabId}`);
            if (activeBtn) {
                activeBtn.classList.remove('text-slate-400', 'hover:bg-slate-800');
                activeBtn.classList.add('bg-blue-600', 'text-white');
            }
            document.querySelectorAll('.tab-content').forEach(content => {
                content.classList.add('hidden');
                content.classList.remove('block');
            });
            document.getElementById(`tab-${tabId}`).classList.remove('hidden');
            document.getElementById(`tab-${tabId}`).classList.add('block');
            
            if (tabId === 'dashboard') { renderTable(); updateStats(); checkLowStockLevels(); }
            else if (tabId === 'material-out') updateMaterialOutSelect();
            else if (tabId === 'dispatch-history') renderDispatchHistory();
            else if (tabId === 'approvals') renderApprovals();
            else if (tabId === 'user-management') renderUsersTable();
        }

        function renderUsersTable() {
            const tbody = document.getElementById('users-table-body');
            tbody.innerHTML = '';
            usersList.forEach(user => {
                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td class="px-6 py-4 font-semibold">${user.username}</td>
                    <td class="px-6 py-4">${user.role}</td>
                    <td class="px-6 py-4 text-xs">${user.created || 'N/A'}</td>
                    <td class="px-6 py-4 text-right">
                        <button onclick="deleteUserAccount('${user.id}')" class="text-rose-600 text-xs font-semibold">Delete</button>
                    </td>`;
                tbody.appendChild(tr);
            });
        }

        function handleCreateUser(e) {
            e.preventDefault();
            const username = document.getElementById('new-user-name').value.trim();
            const password = document.getElementById('new-user-pass').value.trim();
            const role = document.getElementById('new-user-role').value;
            usersList.push({ id: 'usr-' + Date.now(), username, password, role, created: new Date().toISOString().split('T')[0] });
            saveUsers(); renderUsersTable();
            document.getElementById('form-create-user').reset();
            showToast("User created successfully!");
        }

        function deleteUserAccount(id) {
            usersList = usersList.filter(u => u.id !== id);
            saveUsers(); renderUsersTable();
            showToast("User deleted.");
        }

        function toggleDispatchMode() {
            const mode = document.querySelector('input[name="dispatch-mode"]:checked').value;
            const courierFields = document.getElementById('courier-fields');
            if (mode === 'Courier') courierFields.classList.remove('hidden');
            else courierFields.classList.add('hidden');
        }

        function getCategoryFromName(name) {
            for (const [category, models] of Object.entries(catalog)) {
                if (models.includes(name)) return category;
            }
            return 'Other';
        }

        function renderTable() {
            const tbody = document.getElementById('inventory-table-body');
            const searchTerms = document.getElementById('search-input').value.toLowerCase();
            tbody.innerHTML = '';
            const filtered = inventory.filter(item => 
                item.name.toLowerCase().includes(searchTerms) || item.sn.toLowerCase().includes(searchTerms)
            );
            if (filtered.length === 0) {
                tbody.innerHTML = `<tr><td colspan="5" class="px-6 py-12 text-center text-slate-500">No equipment found.</td></tr>`;
                return;
            }
            filtered.forEach(item => {
                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td class="px-6 py-4"><span class="font-semibold">${item.name}</span></td>
                    <td class="px-6 py-4 font-mono">${item.sn}</td>
                    <td class="px-6 py-4"><span class="px-2 py-1 rounded text-xs font-bold ${item.status === 'In' ? 'bg-emerald-100 text-emerald-800' : 'bg-amber-100 text-amber-800'}">${item.status}</span></td>
                    <td class="px-6 py-4 text-xs">${item.status === 'In' ? 'In Stock' : 'Client: ' + item.company}</td>
                    <td class="px-6 py-4 text-right">
                        <button onclick="deleteItem('${item.id}')" class="text-rose-600 text-xs">Delete</button>
                    </td>`;
                tbody.appendChild(tr);
            });
        }

        function renderDispatchHistory() {
            const tbody = document.getElementById('dispatch-history-table-body');
            tbody.innerHTML = '';
            inventory.filter(i => i.status === 'Out').forEach(item => {
                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td class="px-6 py-4 font-semibold">${item.name} (${item.sn})</td>
                    <td class="px-6 py-4">${item.company}</td>
                    <td class="px-6 py-4">${item.dispatchMode}</td>
                    <td class="px-6 py-4">${formatDate(item.dateOut)}</td>
                    <td class="px-6 py-4 text-right"><button onclick="returnToStock('${item.id}')" class="text-xs text-blue-600">Return</button></td>`;
                tbody.appendChild(tr);
            });
        }

        function updateStats() {
            document.getElementById('stat-total').textContent = inventory.length;
            document.getElementById('stat-in').textContent = inventory.filter(i => i.status === 'In').length;
            document.getElementById('stat-out').textContent = inventory.filter(i => i.status === 'Out').length;
        }

        function updateMaterialOutSelect() {
            const select = document.getElementById('out-item');
            select.innerHTML = '<option value="" disabled selected>Select equipment...</option>';
            inventory.filter(i => i.status === 'In').forEach(item => {
                const opt = document.createElement('option');
                opt.value = item.id;
                opt.textContent = `${item.name} (S/N: ${item.sn})`;
                select.appendChild(opt);
            });
        }

        function formatDate(dateString) {
            if (!dateString) return '';
            return new Date(dateString).toLocaleDateString('en-GB');
        }

        document.getElementById('form-material-in').addEventListener('submit', function(e) {
            e.preventDefault();
            inventory.push({
                id: 'INV-' + Date.now(),
                name: document.getElementById('in-name').value,
                sn: document.getElementById('in-sn').value.trim().toUpperCase(),
                source: document.getElementById('in-source').value,
                dateIn: document.getElementById('in-date').value,
                status: 'In',
                timestamp: Date.now()
            });
            saveData(); showToast("Equipment added!");
            this.reset(); document.getElementById('in-date').valueAsDate = new Date();
            switchTab('dashboard');
        });

        document.getElementById('form-material-out').addEventListener('submit', function(e) {
            e.preventDefault();
            const itemId = document.getElementById('out-item').value;
            const item = inventory.find(i => i.id === itemId);
            if (item) {
                item.status = 'Pending Dispatch';
                item.company = document.getElementById('out-company').value;
                item.dispatchMode = document.querySelector('input[name="dispatch-mode"]:checked').value;
                item.dateOut = new Date().toISOString();
                saveData(); showToast("Dispatch requested!");
                this.reset(); switchTab('dashboard');
            }
        });

        function deleteItem(id) {
            inventory = inventory.filter(i => i.id !== id);
            saveData(); renderTable(); updateStats();
            showToast("Item deleted.");
        }

        function returnToStock(id) {
            const item = inventory.find(i => i.id === id);
            if (item) { item.status = 'In'; item.dateOut = null; saveData(); renderDispatchHistory(); updateStats(); showToast("Returned to stock."); }
        }

        function renderApprovals() {
            const tbody = document.getElementById('approvals-table-body');
            tbody.innerHTML = '';
            inventory.filter(i => i.status === 'Pending Dispatch').forEach(item => {
                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td class="px-6 py-4">${item.name} (${item.sn})</td>
                    <td class="px-6 py-4">${item.company}</td>
                    <td class="px-6 py-4 text-right">
                        <button onclick="approveDispatch('${item.id}')" class="px-3 py-1 bg-emerald-100 text-emerald-700 rounded text-xs font-semibold mr-1">Approve</button>
                    </td>`;
                tbody.appendChild(tr);
            });
        }

        function approveDispatch(id) {
            const item = inventory.find(i => i.id === id);
            if (item) { item.status = 'Out'; saveData(); renderApprovals(); updateStats(); showToast("Approved!"); }
        }

        window.addEventListener('DOMContentLoaded', checkAuth);
    </script>
</body>
</html>
