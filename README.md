JANAK
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Janak Positioning & Surveying - Stock Management</title>
    
    <!-- Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>
    
    <!-- Google Fonts -->
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    
    <!-- Phosphor Icons -->
    <script src="https://unpkg.com/@phosphor-icons/web"></script>

    <!-- SheetJS for Excel (.xlsx) Export -->
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    
    <style>
        body {
            font-family: 'Inter', sans-serif;
            background-color: #f8fafc;
        }
        
        /* Custom Scrollbar for the Main Content Area */
        ::-webkit-scrollbar { width: 8px; height: 8px; }
        ::-webkit-scrollbar-track { background: #f1f5f9; border-radius: 4px; }
        ::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 4px; }
        ::-webkit-scrollbar-thumb:hover { background: #94a3b8; }

        /* Print Styles */
        @media print {
            body { background-color: white; }
            #app-container, #toast-container, .no-print, #edit-modal, #camera-modal, #approval-details-modal, #gatepass-preview-modal, #csv-import-modal, #track-modal, #dispatch-details-modal, #manage-models-modal, #drive-sync-modal, #model-pin-modal, #user-edit-modal, #forgot-password-modal {
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
<body class="text-slate-800 antialiased h-screen overflow-hidden flex bg-slate-50">

    <!-- Login Screen -->
    <div id="login-screen" class="fixed inset-0 z-[100] flex items-center justify-center bg-slate-900 px-4">
        <div class="bg-white p-8 rounded-2xl shadow-2xl max-w-sm w-full">
            <div class="flex justify-center mb-5">
                <img src="Janak_Logo.png" alt="Janak Positioning & Surveying Systems Pvt Ltd" class="h-16 w-auto object-contain">
            </div>
            <h2 class="text-2xl font-bold text-center text-slate-800 mb-2">Janak Stock App</h2>
            <p class="text-center text-slate-500 text-sm mb-6">Sign in to your account</p>
            
            <div id="login-error" class="hidden mb-4 p-3 bg-rose-50 text-rose-600 text-sm rounded-lg border border-rose-200 text-center font-medium">
                Invalid username or password
            </div>

            <!-- Quick Role Select Buttons with Highlight Support -->
            <div class="mb-5">
                <p class="text-xs font-semibold text-slate-400 uppercase tracking-wider text-center mb-2.5">Select Role / Quick Login</p>
                <div class="grid grid-cols-2 gap-2.5">
                    <button type="button" id="role-btn-admin" onclick="selectLoginRole('admin')" class="py-2.5 px-3 bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-200 text-xs font-bold rounded-xl transition flex items-center justify-center gap-1.5 shadow-sm">
                        <i class="ph-fill ph-shield-check text-slate-500 text-base" id="role-icon-admin"></i> Admin
                    </button>
                    <button type="button" id="role-btn-manager" onclick="selectLoginRole('manager')" class="py-2.5 px-3 bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-200 text-xs font-bold rounded-xl transition flex items-center justify-center gap-1.5 shadow-sm">
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

    <!-- App Container -->
    <div id="app-container" class="hidden w-full h-full flex-col md:flex-row relative">
        
        <!-- Sidebar -->
        <aside class="w-full md:w-64 bg-slate-900 text-white flex flex-col shadow-xl z-20 shrink-0">
            <div class="p-6 border-b border-slate-800 flex items-center gap-4 md:block">
                <div class="bg-white p-2.5 rounded-xl inline-flex items-center justify-center md:mb-3 shrink-0 shadow-inner">
                    <img src="Janak_Logo.png" alt="Janak Positioning & Surveying" class="h-10 w-auto object-contain">
                </div>
                <div>
                    <h1 class="font-bold text-base leading-tight hidden md:block">Janak Positioning & <br><span class="text-blue-400 font-medium text-xs">Surveying Systems</span></h1>
                    <h1 class="font-bold text-lg leading-tight md:hidden">Janak Stock App</h1>
                </div>
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
        <main class="flex-1 flex flex-col h-full overflow-hidden relative bg-slate-50">
            
            <!-- Header -->
            <header class="h-16 bg-white border-b border-slate-200 flex items-center justify-between px-4 md:px-8 z-10 shrink-0">
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

            <!-- Scrollable Viewport Area (Where the content actually scrolls) -->
            <div id="main-scroll-area" class="flex-1 overflow-y-auto overflow-x-hidden p-4 md:p-8 relative">
                
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
                            
                            <!-- Courier Specific Fields (Hidden by default) -->
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

    <!-- Admin Manage Models & Brands Modal (Inline Editing) -->
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

                <!-- Existing Models List & Inline Renaming -->
                <div>
                    <h4 class="font-bold text-slate-800 text-sm mb-3">Existing Models (Edit name inline and click 'Save' to update database & Drive without losing old data)</h4>
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
                    <p class="text-xs text-slate-500 mt-1">Expected format: Model names in row 1 (e.g. <code class="bg-slate-100 px-1 py-0.5 rounded text-slate-700 font-mono">A90, A66</code>), with serial numbers listed vertically below each column.</p>
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
                <!-- Timeline content populated by JS -->
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
                    <textarea id="approval-notes" rows="3" required placeholder="e.g. Verified client purchase order, checked physical item condition, approved by Director." 
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

    <!-- Printable Gatepass View (With Embedded Logo) -->
    <div id="gatepass-print-view" class="bg-white text-black font-sans">
        <div class="text-center mb-6 border-b-2 border-black pb-4">
            <div class="flex justify-center mb-3">
                <img src="Janak_Logo.png" alt="Janak Positioning & Surveying Systems Pvt Ltd" class="h-16 w-auto object-contain">
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
                "JAVAD - TRIUMPH-3NR",
                "JAVAD - TRIUMPH-1M",
                "JAVAD - TRIUMPH-1M Plus",
                "JAVAD - TRIUMPH-2",
                "JAVAD - TRIUMPH-3",
                "JAVAD - TRIUMPH-LS Plus",
                "JAVAD - PulserOne"
            ],
            "JAVAD - OEM GNSS Boards": [
                "JAVAD - TR-2S OEM Board",
                "JAVAD - TRE-3S OEM Board",
                "JAVAD - TRH Heading OEM Board"
            ],
            "JAVAD - GNSS Antennas": [
                "JAVAD - GrAnt-G3T-JS",
                "JAVAD - GrAnt-G3-JS",
                "JAVAD - GrAnt-G2T-JSM",
                "JAVAD - GrAnt-3L",
                "JAVAD - GrAnt-3LS",
                "JAVAD - SpaceAnt",
                "JAVAD - MiniAnt-G3",
                "JAVAD - MiniAnt-G3T"
            ],
            "JAVAD - INS Solutions": [
                "JAVAD - ALPHA INS",
                "JAVAD - TR-2S INS",
                "JAVAD - TR-2S AJ INS",
                "JAVAD - TR-2S HDA INS",
                "JAVAD - TR-3S DUO INS"
            ],
            "FOIF Brand": [
                "FOIF - A90",
                "FOIF - A66",
                "FOIF - RTS 102",
                "FOIF - RTS 392",
                "FOIF - RTS350"
            ],
            "Alpha Geo Brand": [
                "Alpha Geo - GNSS Receiver",
                "Alpha Geo - Slam"
            ],
            "Janak Brand": [
                "Janak Brand - GNSS Receiver",
                "Janak Brand - Controller"
            ]
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

        // Highlight and fill quick login role buttons
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

        // --- Forgot Password Recovery Functions ---
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
                showToast("Admin password successfully reset! Please login with your new credentials.");
                
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
                if(!document.getElementById('tab-approvals').classList.contains('hidden') || !document.getElementById('tab-user-management').classList.contains('hidden')) {
                    switchTab('dashboard');
                }
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
            if (data) {
                try {
                    catalog = JSON.parse(data);
                } catch (e) {
                    catalog = JSON.parse(JSON.stringify(DEFAULT_CATALOG));
                }
            } else {
                catalog = JSON.parse(JSON.stringify(DEFAULT_CATALOG));
                localStorage.setItem(CATALOG_KEY, JSON.stringify(catalog));
            }
        }

        function saveCatalog() {
            localStorage.setItem(CATALOG_KEY, JSON.stringify(catalog));
            syncToGoogleDrive(false);
        }

        function loadUsers() {
            const data = localStorage.getItem(USERS_KEY);
            if (data) {
                try {
                    usersList = JSON.parse(data);
                } catch (e) {
                    usersList = JSON.parse(JSON.stringify(DEFAULT_USERS));
                }
            } else {
                usersList = JSON.parse(JSON.stringify(DEFAULT_USERS));
                localStorage.setItem(USERS_KEY, JSON.stringify(usersList));
            }
        }

        function saveUsers() {
            localStorage.setItem(USERS_KEY, JSON.stringify(usersList));
            syncToGoogleDrive(false);
        }

        function loadData() {
            const data = localStorage.getItem(STORAGE_KEY);
            if (data) {
                try {
                    inventory = JSON.parse(data);
                } catch (e) {
                    inventory = [];
                }
            } else {
                inventory = [];
            }
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
            const enteredPin = document.getElementById('model-pin-input').value.trim();
            if (enteredPin === '8317') {
                closeModelPinModal();
                openManageModelsModal();
            } else {
                document.getElementById('pin-error').classList.remove('hidden');
                document.getElementById('model-pin-input').value = '';
                document.getElementById('model-pin-input').focus();
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
            } else if (state === 'error') {
                badge.className = "px-2.5 py-1 rounded-full text-[11px] font-bold border flex items-center gap-1.5 transition bg-rose-50 text-rose-700 border-rose-200";
                dot.className = "w-2 h-2 rounded-full bg-rose-500";
                text.textContent = "Drive Sync Error";
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
            if (folderLink) {
                window.open(folderLink, '_blank');
            } else {
                showToast("Please paste your Google Drive folder link in the reference box first.", "info");
            }
        }

        function saveDriveWebhookUrl() {
            const url = document.getElementById('drive-webhook-url').value.trim();
            const folderUrl = document.getElementById('drive-folder-url').value.trim();
            localStorage.setItem(DRIVE_WEBHOOK_KEY, url);
            localStorage.setItem(DRIVE_FOLDER_LINK_KEY, folderUrl);
            updateDriveSyncBadge();
            showToast("Google Drive Sync settings saved!");
            if (url) {
                syncToGoogleDrive(true);
            }
        }

        function checkAutoDailyDriveBackup() {
            const webhookUrl = localStorage.getItem(DRIVE_WEBHOOK_KEY);
            if (!webhookUrl || inventory.length === 0) return;

            const todayStr = new Date().toISOString().split('T')[0];
            const lastSync = localStorage.getItem(LAST_SYNC_DATE_KEY);

            if (!lastSync || !lastSync.startsWith(todayStr)) {
                syncToGoogleDrive(false);
            }
        }

        async function syncToGoogleDrive(manual = false) {
            const webhookUrl = localStorage.getItem(DRIVE_WEBHOOK_KEY);
            if (!webhookUrl) {
                if (manual) showToast("Please paste your Google Apps Script Web App URL first.", "error");
                return;
            }

            updateDriveSyncBadge('syncing');
            if (manual) showToast("Pushing data & Excel sheet to Google Drive...", "info");

            try {
                await fetch(webhookUrl, {
                    method: 'POST',
                    mode: 'no-cors',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        catalog: catalog,
                        users: usersList,
                        inventory: inventory,
                        timestamp: new Date().toISOString()
                    })
                });

                const nowStr = new Date().toISOString().split('T')[0] + ' ' + new Date().toLocaleTimeString();
                localStorage.setItem(LAST_SYNC_DATE_KEY, nowStr);
                const displayEl = document.getElementById('last-drive-sync-display');
                if (displayEl) displayEl.textContent = nowStr;

                updateDriveSyncBadge('idle');
                if (manual) showToast("Synchronized with Google Drive successfully!");
            } catch (err) {
                updateDriveSyncBadge('error');
                if (manual) showToast("Failed to reach Google Drive webhook.", "error");
            }
        }

        async function loadFromGoogleDrive(manual = false) {
            const webhookUrl = localStorage.getItem(DRIVE_WEBHOOK_KEY);
            if (!webhookUrl) {
                if (manual) showToast("Please paste your Google Apps Script Web App URL first.", "error");
                return;
            }

            updateDriveSyncBadge('syncing');
            if (manual) showToast("Fetching latest database from Google Drive...", "info");

            try {
                const response = await fetch(webhookUrl, { method: 'GET' });
                const data = await response.json();

                if (data && data.status === 'success') {
                    if (data.catalog && typeof data.catalog === 'object') {
                        catalog = data.catalog;
                        localStorage.setItem(CATALOG_KEY, JSON.stringify(catalog));
                    }
                    if (data.users && Array.isArray(data.users)) {
                        usersList = data.users;
                        localStorage.setItem(USERS_KEY, JSON.stringify(usersList));
                    }
                    if (Array.isArray(data.inventory) && data.inventory.length > 0) {
                        data.inventory.forEach(remoteItem => {
                            const localIdx = inventory.findIndex(i => i.id === remoteItem.id);
                            if (localIdx > -1) {
                                inventory[localIdx] = remoteItem;
                            } else {
                                inventory.push(remoteItem);
                            }
                        });
                        localStorage.setItem(STORAGE_KEY, JSON.stringify(inventory));
                    }

                    populateModelDropdowns();
                    renderTable();
                    renderUsersTable();
                    updateStats();
                    updateMaterialOutSelect();
                    checkLowStockLevels();

                    const nowStr = new Date().toISOString().split('T')[0] + ' ' + new Date().toLocaleTimeString();
                    localStorage.setItem(LAST_SYNC_DATE_KEY, nowStr);
                    const displayEl = document.getElementById('last-drive-sync-display');
                    if (displayEl) displayEl.textContent = nowStr;

                    updateDriveSyncBadge('idle');
                    if (manual) showToast("Data pulled and synchronized from Google Drive!");
                } else {
                    updateDriveSyncBadge('idle');
                    if (manual) showToast("No existing backup found in Drive yet. Click 'Push to Drive' first.", "info");
                }
            } catch (err) {
                updateDriveSyncBadge('idle');
                if (manual) showToast("Could not pull from Drive. Ensure Web App permissions are set to 'Anyone'.", "error");
            }
        }

        function exportToExcel() {
            if (inventory.length === 0) {
                showToast("No inventory data to export", "error");
                return;
            }

            const dateStr = new Date().toISOString().split('T')[0];

            const masterData = inventory.map(item => ({
                "Item ID": item.id || "",
                "Category": getCategoryFromName(item.name),
                "Equipment Model": item.name || "",
                "Serial Number": item.sn || "",
                "Status": item.status || "",
                "Received From (Source)": item.source || "",
                "Date Received": item.dateIn || "",
                "Accessories": item.accessories || "",
                "Dispatch Mode": item.dispatchMode || "",
                "Courier Partner": item.courierName || "",
                "AWB / Tracking No": item.awb || "",
                "Client / Company": item.company || "",
                "Challan / Invoice No": item.challan || "",
                "Contact Phone": item.phone || "",
                "Delivery Location": item.dispatchTo || "",
                "Date Dispatched": item.dateOut ? formatDate(item.dateOut) : "",
                "Remarks": item.remarks || "",
                "Admin Approval Notes": item.approvalNotes || ""
            }));

            const summaryRows = [];
            for (const [cat, models] of Object.entries(catalog)) {
                models.forEach(model => {
                    const inCount = inventory.filter(i => i.name === model && i.status === 'In').length;
                    const outCount = inventory.filter(i => i.name === model && i.status === 'Out').length;
                    summaryRows.push({
                        "Brand / Category": cat,
                        "Equipment Model": model,
                        "In Stock": inCount,
                        "Dispatched (Out)": outCount,
                        "Total Units": inCount + outCount
                    });
                });
            }

            const wb = XLSX.utils.book_new();
            const wsMaster = XLSX.utils.json_to_sheet(masterData);
            const wsSummary = XLSX.utils.json_to_sheet(summaryRows);

            XLSX.utils.book_append_sheet(wb, wsMaster, "Master_Inventory");
            XLSX.utils.book_append_sheet(wb, wsSummary, "Stock_Summary");

            XLSX.writeFile(wb, `Janak_Stock_Backup_${dateStr}.xlsx`);
            showToast("Excel (.xlsx) backup downloaded!");
        }

        function populateModelDropdowns() {
            const selects = [document.getElementById('in-name'), document.getElementById('edit-name')];
            
            selects.forEach(select => {
                if (!select) return;
                const currentVal = select.value;
                select.innerHTML = '<option value="" disabled selected>Select categorized equipment model...</option>';

                for (const [category, models] of Object.entries(catalog)) {
                    const optgroup = document.createElement('optgroup');
                    optgroup.label = category;

                    models.forEach(model => {
                        const option = document.createElement('option');
                        option.value = model;
                        option.textContent = model;
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
                    opt.value = cat;
                    opt.textContent = cat;
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
            populateModelDropdowns();
            renderTable();
            updateStats();
            updateMaterialOutSelect();
        }

        function renderManageModelsList() {
            const container = document.getElementById('models-management-list');
            container.innerHTML = '';

            const categories = Object.keys(catalog);
            categories.forEach((category, catIdx) => {
                const models = catalog[category] || [];
                let modelsHtml = '';

                models.forEach((model, modIdx) => {
                    const inputId = `model-input-${catIdx}-${modIdx}`;
                    const safeModelVal = model.replace(/"/g, '&quot;');
                    modelsHtml += `
                        <div class="flex items-center justify-between gap-2 bg-white p-2 rounded-lg border border-slate-200 text-xs mb-1.5">
                            <input type="text" id="${inputId}" value="${safeModelVal}" class="flex-1 font-mono font-medium text-slate-800 px-2.5 py-1.5 border border-slate-200 rounded-md focus:border-blue-500 focus:ring-1 focus:ring-blue-500 outline-none">
                            <div class="flex items-center gap-1 shrink-0">
                                <button type="button" onclick="saveInlineModelRename(${catIdx}, ${modIdx}, '${inputId}')" class="px-3 py-1.5 bg-blue-600 hover:bg-blue-700 text-white rounded-md font-semibold transition">Save</button>
                                <button type="button" onclick="deleteCatalogModel(${catIdx}, ${modIdx})" class="px-2.5 py-1.5 bg-rose-50 hover:bg-rose-100 text-rose-600 rounded-md font-semibold transition" title="Remove Model">Delete</button>
                            </div>
                        </div>
                    `;
                });

                const div = document.createElement('div');
                div.className = 'bg-slate-50 p-3.5 rounded-xl border border-slate-200';
                div.innerHTML = `
                    <div class="flex justify-between items-center mb-2 pb-1 border-b border-blue-200">
                        <h5 class="font-bold text-xs text-blue-700 uppercase">${category}</h5>
                        <span class="text-[11px] text-slate-400 font-semibold">${models.length} model(s)</span>
                    </div>
                    <div class="max-h-48 overflow-y-auto space-y-1 pr-1">${modelsHtml || '<p class="text-xs text-slate-400 italic py-1">No models in this brand group yet.</p>'}</div>
                `;
                container.appendChild(div);
            });
        }

        function handleAddNewCategory(e) {
            e.preventDefault();
            const catInput = document.getElementById('new-category-name');
            const newCat = catInput.value.trim();
            if (!newCat) return;

            if (catalog[newCat]) {
                showToast("Brand / Category already exists!", "error");
                return;
            }

            catalog[newCat] = [];
            saveCatalog();
            catInput.value = '';
            populateModelDropdowns();
            renderManageModelsList();
            showToast(`Brand group "${newCat}" created & synced!`);
        }

        function handleAddNewModel(e) {
            e.preventDefault();
            const category = document.getElementById('new-model-category').value;
            const modelName = document.getElementById('new-model-name').value.trim();

            if (!catalog[category]) {
                catalog[category] = [];
            }

            if (catalog[category].includes(modelName)) {
                showToast("Model already exists in this category!", "error");
                return;
            }

            catalog[category].push(modelName);
            saveCatalog();
            document.getElementById('new-model-name').value = '';
            populateModelDropdowns();
            renderManageModelsList();
            updateStats();
            showToast(`Added "${modelName}" to ${category} & synced!`);
        }

        function saveInlineModelRename(catIdx, modIdx, inputId) {
            const categories = Object.keys(catalog);
            const category = categories[catIdx];
            const oldName = catalog[category][modIdx];
            const inputEl = document.getElementById(inputId);
            if (!inputEl) return;

            const trimmedNewName = inputEl.value.trim();
            if (!trimmedNewName) {
                showToast("Model name cannot be empty!", "error");
                inputEl.value = oldName;
                return;
            }

            if (trimmedNewName === oldName) {
                showToast("No changes made to model name.", "info");
                return;
            }

            if (catalog[category].includes(trimmedNewName)) {
                showToast("A model with this name already exists in this category!", "error");
                return;
            }

            catalog[category][modIdx] = trimmedNewName;
            localStorage.setItem(CATALOG_KEY, JSON.stringify(catalog));

            let updatedCount = 0;
            inventory.forEach(item => {
                if (item.name === oldName) {
                    item.name = trimmedNewName;
                    updatedCount++;
                }
            });
            saveData();

            populateModelDropdowns();
            renderManageModelsList();
            renderTable();
            updateStats();
            updateMaterialOutSelect();
            showToast(`Renamed to "${trimmedNewName}" & synced ${updatedCount} record(s) to Drive!`);
        }

        function deleteCatalogModel(catIdx, modIdx) {
            const categories = Object.keys(catalog);
            const category = categories[catIdx];
            const modelName = catalog[category][modIdx];

            catalog[category].splice(modIdx, 1);
            saveCatalog();
            populateModelDropdowns();
            renderManageModelsList();
            updateStats();
            showToast(`Removed "${modelName}" from catalog list.`);
        }

        function checkLowStockLevels() {
            const modelCounts = {};
            inventory.forEach(item => {
                if (item.status === 'In') {
                    modelCounts[item.name] = (modelCounts[item.name] || 0) + 1;
                }
            });

            const widgetContainer = document.getElementById('low-stock-widget-container');
            const lowStockList = document.getElementById('low-stock-list');
            lowStockList.innerHTML = '';
            
            let lowStockModels = [];

            for (const [model, count] of Object.entries(modelCounts)) {
                if (count <= MIN_STOCK_THRESHOLD) {
                    lowStockModels.push({ model, count });
                    
                    const alertKey = `low_stock_alert_${model}`;
                    if (!sessionStorage.getItem(alertKey)) {
                        sessionStorage.setItem(alertKey, 'true');
                        triggerLowStockEmail(model, count);
                    }
                }
            }

            if (lowStockModels.length > 0) {
                widgetContainer.classList.remove('hidden');
                lowStockModels.forEach(item => {
                    const div = document.createElement('div');
                    div.className = 'bg-white p-3 rounded-xl border border-rose-200 shadow-sm flex justify-between items-center';
                    div.innerHTML = `
                        <span class="font-semibold text-slate-800 text-xs">${item.model}</span>
                        <span class="px-2 py-0.5 bg-rose-100 text-rose-700 rounded text-xs font-bold">${item.count} left</span>
                    `;
                    lowStockList.appendChild(div);
                });
            } else {
                widgetContainer.classList.add('hidden');
            }
        }

        function triggerLowStockEmail(modelName, currentCount) {
            const subject = encodeURIComponent(`LOW STOCK ALERT: ${modelName} (${currentCount} remaining)`);
            const body = encodeURIComponent(`Dear Admin,\n\nThis is an automated inventory alert from Janak Stock Management System.\n\nThe stock count for model "${modelName}" has dropped to ${currentCount} unit(s), which is at or below the minimum required threshold of ${MIN_STOCK_THRESHOLD}.\n\nPlease arrange for restocking promptly.\n\nBest regards,\nJanak Stock System`);
            
            const mailtoLink = `mailto:${ADMIN_EMAIL}?subject=${subject}&body=${body}`;
            window.location.href = mailtoLink;
            showToast(`Low stock warning for ${modelName}! Email triggered to ${ADMIN_EMAIL}`, "error");
        }

        function showToast(message, type = 'success') {
            const container = document.getElementById('toast-container');
            const toast = document.createElement('div');
            toast.className = `flex items-center gap-3 px-5 py-3 rounded-xl shadow-lg border text-sm font-medium transform transition-all duration-300 translate-y-10 opacity-0`;
            
            if (type === 'success') {
                toast.classList.add('bg-slate-900', 'text-white', 'border-slate-800');
                toast.innerHTML = `<i class="ph-fill ph-check-circle text-emerald-400 text-xl"></i> ${message}`;
            } else if (type === 'error') {
                toast.classList.add('bg-white', 'text-rose-600', 'border-rose-200');
                toast.innerHTML = `<i class="ph-fill ph-warning-circle text-rose-500 text-xl"></i> ${message}`;
            } else {
                toast.classList.add('bg-white', 'text-blue-600', 'border-blue-200');
                toast.innerHTML = `<i class="ph-fill ph-info text-blue-500 text-xl"></i> ${message}`;
            }

            container.appendChild(toast);
            requestAnimationFrame(() => toast.classList.remove('translate-y-10', 'opacity-0'));

            setTimeout(() => {
                toast.classList.add('opacity-0');
                setTimeout(() => toast.remove(), 300);
            }, 3000);
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
            if(activeBtn) {
                activeBtn.classList.remove('text-slate-400', 'hover:bg-slate-800');
                activeBtn.classList.add('bg-blue-600', 'text-white');
            }

            document.querySelectorAll('.tab-content').forEach(content => {
                content.classList.add('hidden');
                content.classList.remove('block');
            });
            document.getElementById(`tab-${tabId}`).classList.remove('hidden');
            document.getElementById(`tab-${tabId}`).classList.add('block');

            const titles = {
                'dashboard': 'Inventory Dashboard',
                'material-in': 'Receive Material',
                'material-out': 'Dispatch Material',
                'dispatch-history': 'Dispatch History Log',
                'approvals': 'Dispatch Approvals',
                'user-management': 'User Accounts Management'
            };
            document.getElementById('header-title').textContent = titles[tabId];
            
            if(tabId === 'dashboard') {
                renderTable();
                updateStats();
                checkLowStockLevels();
            } else if (tabId === 'material-in') {
                populateModelDropdowns();
            } else if (tabId === 'material-out') {
                updateMaterialOutSelect();
            } else if (tabId === 'dispatch-history') {
                renderDispatchHistory();
            } else if (tabId === 'approvals') {
                renderApprovals();
            } else if (tabId === 'user-management') {
                renderUsersTable();
            }
        }

        // --- User Management Functions ---
        function renderUsersTable() {
            const tbody = document.getElementById('users-table-body');
            const countBadge = document.getElementById('user-count-badge');
            tbody.innerHTML = '';
            
            loadUsers();
            countBadge.textContent = `${usersList.length} User(s)`;

            if (usersList.length === 0) {
                tbody.innerHTML = `<tr><td colspan="4" class="px-6 py-12 text-center text-slate-500 bg-slate-50">No user accounts created yet.</td></tr>`;
                return;
            }

            usersList.forEach(user => {
                const tr = document.createElement('tr');
                tr.className = 'hover:bg-slate-50 transition-colors';

                const roleBadge = user.role === 'admin'
                    ? `<span class="px-2.5 py-1 rounded-md text-xs font-bold bg-purple-100 text-purple-800 border border-purple-200">Admin</span>`
                    : `<span class="px-2.5 py-1 rounded-md text-xs font-bold bg-blue-100 text-blue-800 border border-blue-200">Store Manager</span>`;

                tr.innerHTML = `
                    <td class="px-6 py-4 font-semibold text-slate-900">${user.username}</td>
                    <td class="px-6 py-4">${roleBadge}</td>
                    <td class="px-6 py-4 text-xs text-slate-500">${user.created || 'N/A'}</td>
                    <td class="px-6 py-4 text-right space-x-2">
                        <button type="button" onclick="openUserEditModal('${user.id}')" class="px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg text-xs font-semibold transition" title="Edit User">
                            <i class="ph ph-pencil-simple"></i> Edit
                        </button>
                        ${user.username !== currentUsername && user.username !== 'admin' ? `
                        <button type="button" onclick="deleteUserAccount('${user.id}')" class="px-3 py-1.5 bg-rose-50 hover:bg-rose-100 text-rose-600 rounded-lg text-xs font-semibold transition" title="Delete User">
                            <i class="ph ph-trash"></i> Delete
                        </button>` : `<span class="text-xs text-slate-400 italic">Protected</span>`}
                    </td>
                `;
                tbody.appendChild(tr);
            });
        }

        function handleCreateUser(e) {
            e.preventDefault();
            if (currentUserRole !== 'admin') {
                showToast("Permission denied. Admins only.", "error");
                return;
            }

            const username = document.getElementById('new-user-name').value.trim();
            const password = document.getElementById('new-user-pass').value.trim();
            const role = document.getElementById('new-user-role').value;

            if (!username || !password) {
                showToast("Please enter both username and password.", "error");
                return;
            }

            const existing = usersList.find(u => u.username.toLowerCase() === username.toLowerCase());
            if (existing) {
                showToast("Username already exists. Choose a different name.", "error");
                return;
            }

            const newUser = {
                id: 'usr-' + Date.now().toString(),
                username: username,
                password: password,
                role: role,
                created: new Date().toISOString().split('T')[0]
            };

            usersList.push(newUser);
            saveUsers();
            renderUsersTable();
            document.getElementById('form-create-user').reset();
            showToast(`User account "${username}" created successfully & synced!`);
        }

        function openUserEditModal(id) {
            const user = usersList.find(u => u.id === id);
            if (!user) return;

            document.getElementById('edit-user-id').value = user.id;
            document.getElementById('edit-user-name').value = user.username;
            document.getElementById('edit-user-pass').value = '';
            document.getElementById('edit-user-role').value = user.role;

            document.getElementById('user-edit-modal').classList.remove('hidden');
        }

        function closeUserEditModal() {
            document.getElementById('user-edit-modal').classList.add('hidden');
        }

        function handleSaveEditedUser(e) {
            e.preventDefault();
            if (currentUserRole !== 'admin') {
                showToast("Permission denied.", "error");
                return;
            }

            const id = document.getElementById('edit-user-id').value;
            const newUsername = document.getElementById('edit-user-name').value.trim();
            const newPassword = document.getElementById('edit-user-pass').value.trim();
            const newRole = document.getElementById('edit-user-role').value;

            const userIndex = usersList.findIndex(u => u.id === id);
            if (userIndex === -1) return;

            const duplicate = usersList.find(u => u.username.toLowerCase() === newUsername.toLowerCase() && u.id !== id);
            if (duplicate) {
                showToast("Username already taken by another account.", "error");
                return;
            }

            usersList[userIndex].username = newUsername;
            if (newPassword) {
                usersList[userIndex].password = newPassword;
            }
            usersList[userIndex].role = newRole;

            saveUsers();
            renderUsersTable();
            closeUserEditModal();
            showToast(`User account updated & synced!`);
        }

        function deleteUserAccount(id) {
            if (currentUserRole !== 'admin') {
                showToast("Permission denied.", "error");
                return;
            }

            const user = usersList.find(u => u.id === id);
            if (user && (user.username === currentUsername || user.username === 'admin')) {
                showToast("Cannot delete your active session or master admin account.", "error");
                return;
            }

            usersList = usersList.filter(u => u.id !== id);
            saveUsers();
            renderUsersTable();
            showToast("User account deleted.");
        }

        function toggleDispatchMode() {
            const mode = document.querySelector('input[name="dispatch-mode"]:checked').value;
            const courierFields = document.getElementById('courier-fields');
            const courierName = document.getElementById('out-courier-name');

            if (mode === 'Courier') {
                courierFields.classList.remove('hidden');
                courierName.setAttribute('required', 'true');
            } else {
                courierFields.classList.add('hidden');
                courierName.removeAttribute('required');
            }
        }

        function getCategoryFromName(name) {
            for (const [category, models] of Object.entries(catalog)) {
                if (models.includes(name)) return category;
            }
            if (name.startsWith('JAVAD')) return 'JAVAD - Smart Antennas';
            if (name.startsWith('FOIF')) return 'FOIF Brand';
            if (name.startsWith('Alpha Geo')) return 'Alpha Geo Brand';
            if (name.startsWith('Janak Brand')) return 'Janak Brand';
            return 'Other / Uncategorized';
        }

        function renderTable() {
            const tbody = document.getElementById('inventory-table-body');
            const searchTerms = document.getElementById('search-input').value.toLowerCase();
            
            tbody.innerHTML = '';
            const sortedInventory = [...inventory].sort((a, b) => b.timestamp - a.timestamp);
            
            const filtered = sortedInventory.filter(item => {
                const category = getCategoryFromName(item.name).toLowerCase();
                return item.name.toLowerCase().includes(searchTerms) || 
                       category.includes(searchTerms) ||
                       item.sn.toLowerCase().includes(searchTerms) ||
                       (item.source && item.source.toLowerCase().includes(searchTerms)) ||
                       (item.company && item.company.toLowerCase().includes(searchTerms)) ||
                       (item.challan && item.challan.toLowerCase().includes(searchTerms));
            });

            if (filtered.length === 0) {
                tbody.innerHTML = `<tr><td colspan="5" class="px-6 py-12 text-center text-slate-500 bg-slate-50">No equipment found matching your criteria.</td></tr>`;
                return;
            }

            filtered.forEach(item => {
                const tr = document.createElement('tr');
                tr.className = 'hover:bg-slate-50 transition-colors group';

                const isOut = item.status === 'Out';
                const isPending = item.status === 'Pending Dispatch';
                const category = getCategoryFromName(item.name);
                
                let statusBadge = '';
                if (isOut) {
                    statusBadge = `<span class="inline-flex items-center px-2.5 py-1 rounded-md text-xs font-bold bg-amber-100 text-amber-800 border border-amber-200">${item.dispatchMode === 'Courier' ? 'COURIER OUT' : 'DISPATCHED'}</span>`;
                } else if (isPending) {
                    statusBadge = `<span class="inline-flex items-center px-2.5 py-1 rounded-md text-xs font-bold bg-purple-100 text-purple-800 border border-purple-200">PENDING</span>`;
                } else {
                    statusBadge = `<span class="inline-flex items-center px-2.5 py-1 rounded-md text-xs font-bold bg-emerald-100 text-emerald-800 border border-emerald-200">IN STOCK</span>`;
                }

                const activityInfo = isOut || isPending
                    ? `<div class="text-xs"><span class="text-slate-500">${isPending ? 'Req' : 'Out'}:</span> ${formatDate(item.dateOut)}<br><span class="font-medium text-slate-700 truncate block w-48" title="${item.company} (${item.challan || 'N/A'})">Client: ${item.company || 'N/A'} [${item.challan || 'N/A'}]</span></div>`
                    : `<div class="text-xs"><span class="text-slate-500">In:</span> ${formatDate(item.dateIn)}<br><span class="text-slate-400">Accessories: ${item.accessories || 'None'}</span></div>`;

                tr.innerHTML = `
                    <td class="px-6 py-4">
                        <span class="inline-block px-2 py-0.5 mb-1 bg-slate-100 text-slate-600 rounded text-[10px] font-bold uppercase tracking-wider border border-slate-200">${category}</span>
                        <span class="font-semibold text-slate-900 block">${item.name}</span>
                        <span class="text-xs text-slate-500 font-medium">From: ${item.source || 'N/A'}</span>
                    </td>
                    <td class="px-6 py-4"><span class="font-mono text-sm bg-slate-100 text-slate-700 px-2 py-1 rounded border border-slate-200">${item.sn}</span></td>
                    <td class="px-6 py-4">${statusBadge}</td>
                    <td class="px-6 py-4">${activityInfo}</td>
                    <td class="px-6 py-4 text-right">
                        <div class="flex items-center justify-end gap-1.5 opacity-90 group-hover:opacity-100 transition-opacity">
                            <button type="button" onclick="openTrackModal('${item.id}')" class="p-1.5 text-blue-600 hover:bg-blue-50 rounded" title="Track Asset Journey">
                                <i class="ph ph-compass text-lg"></i>
                            </button>
                            ${isOut ? 
                                `<button type="button" onclick="returnToStock('${item.id}')" class="p-1.5 text-emerald-600 hover:bg-emerald-50 rounded" title="Mark as Returned (In Stock)">
                                    <i class="ph ph-arrow-u-down-left text-lg"></i>
                                </button>` : ''
                            }
                            ${currentUserRole === 'admin' ? `
                            <button type="button" onclick="openEditModal('${item.id}')" class="p-1.5 text-slate-500 hover:text-blue-600 hover:bg-slate-100 rounded" title="Edit Item">
                                <i class="ph ph-pencil-simple text-lg"></i>
                            </button>
                            <button type="button" onclick="deleteItem('${item.id}')" class="p-1.5 text-slate-500 hover:text-rose-600 hover:bg-rose-50 rounded" title="Delete Item">
                                <i class="ph ph-trash text-lg"></i>
                            </button>
                            ` : ''}
                        </div>
                    </td>
                `;
                tbody.appendChild(tr);
            });
        }

        function renderDispatchHistory() {
            const tbody = document.getElementById('dispatch-history-table-body');
            tbody.innerHTML = '';

            const dateFrom = document.getElementById('filter-date-from').value;
            const dateTo = document.getElementById('filter-date-to').value;

            let dispatchedItems = inventory.filter(i => i.status === 'Out');

            if (dateFrom) {
                dispatchedItems = dispatchedItems.filter(i => {
                    if (!i.dateOut) return false;
                    const itemDate = i.dateOut.split('T')[0];
                    return itemDate >= dateFrom;
                });
            }
            if (dateTo) {
                dispatchedItems = dispatchedItems.filter(i => {
                    if (!i.dateOut) return false;
                    const itemDate = i.dateOut.split('T')[0];
                    return itemDate <= dateTo;
                });
            }

            const sorted = dispatchedItems.sort((a, b) => new Date(b.dateOut) - new Date(a.dateOut));

            if (sorted.length === 0) {
                tbody.innerHTML = `<tr><td colspan="5" class="px-6 py-12 text-center text-slate-500 bg-slate-50">No dispatch records found for the selected criteria.</td></tr>`;
                return;
            }

            sorted.forEach(item => {
                const tr = document.createElement('tr');
                tr.className = 'hover:bg-slate-50 transition-colors';

                const modeBadge = item.dispatchMode === 'Courier' 
                    ? `<span class="px-2 py-0.5 bg-blue-50 text-blue-700 border border-blue-200 rounded text-xs font-semibold">Courier (${item.courierName})</span><br><span class="text-xs font-mono text-slate-500">AWB: ${item.awb || 'N/A'}</span>`
                    : `<span class="px-2 py-0.5 bg-slate-100 text-slate-700 border border-slate-200 rounded text-xs font-semibold">Direct Delivery</span>`;

                tr.innerHTML = `
                    <td class="px-6 py-4">
                        <span class="font-semibold text-slate-900 block">${item.name}</span>
                        <span class="text-xs font-mono text-slate-500">S/N: ${item.sn}</span>
                    </td>
                    <td class="px-6 py-4">
                        <span class="font-bold text-slate-800 block">${item.company || 'N/A'}</span>
                        <span class="text-xs text-slate-500">Challan: ${item.challan || 'N/A'}</span>
                    </td>
                    <td class="px-6 py-4">${modeBadge}</td>
                    <td class="px-6 py-4 text-sm text-slate-700">${formatDate(item.dateOut)}</td>
                    <td class="px-6 py-4 text-right">
                        <button type="button" onclick="openDispatchDetailsModal('${item.id}')" class="px-3 py-1.5 bg-blue-50 text-blue-700 hover:bg-blue-100 rounded font-semibold text-xs transition mr-1">
                            <i class="ph ph-file-text"></i> Details / Proof
                        </button>
                    </td>
                `;
                tbody.appendChild(tr);
            });
        }

        function resetDispatchFilters() {
            document.getElementById('filter-date-from').value = '';
            document.getElementById('filter-date-to').value = '';
            renderDispatchHistory();
        }

        function openDispatchDetailsModal(id) {
            const item = inventory.find(i => i.id === id);
            if (!item) return;

            const content = document.getElementById('dispatch-details-content');
            content.innerHTML = `
                <div class="bg-slate-50 p-4 rounded-xl border border-slate-200 space-y-2">
                    <h4 class="font-bold text-slate-900 text-base">${item.name}</h4>
                    <p class="text-xs font-mono text-slate-600">Serial Number: <strong>${item.sn}</strong></p>
                    <p class="text-xs text-slate-600">Received Source: <strong>${item.source || 'N/A'}</strong></p>
                    <p class="text-xs text-slate-600">Accessories: <strong>${item.accessories || 'None'}</strong></p>
                </div>
                
                <div class="space-y-2 pt-2">
                    <p class="text-xs font-bold uppercase tracking-wider text-slate-400">Client & Delivery Info</p>
                    <p><strong>Company / Client:</strong> ${item.company || 'N/A'}</p>
                    <p><strong>Challan Number:</strong> ${item.challan || 'N/A'}</p>
                    <p><strong>Delivery Location:</strong> ${item.dispatchTo || 'N/A'}</p>
                    <p><strong>Contact Phone:</strong> ${item.phone || 'N/A'}</p>
                    <p><strong>Dispatch Date:</strong> ${formatDate(item.dateOut)}</p>
                    <p><strong>Dispatch Mode:</strong> ${item.dispatchMode || 'Direct'} ${item.dispatchMode === 'Courier' ? `(Partner: ${item.courierName}, AWB: ${item.awb || 'N/A'})` : ''}</p>
                </div>

                <div class="space-y-2 pt-2 border-t border-slate-200">
                    <p class="text-xs font-bold uppercase tracking-wider text-slate-400">Admin Approval & Verification Proof</p>
                    <p class="italic bg-slate-50 p-3 rounded-lg border border-slate-200 text-slate-700">"${item.approvalNotes || 'No notes recorded.'}"</p>
                    ${item.verificationPhoto ? `
                        <div class="mt-2">
                            <p class="text-xs font-semibold text-slate-600 mb-1">Captured Proof Photo:</p>
                            <img src="${item.verificationPhoto}" class="w-full h-48 object-cover rounded-xl border shadow-sm">
                        </div>
                    ` : '<p class="text-xs text-slate-500 italic">No verification photo attached.</p>'}
                </div>
            `;

            document.getElementById('dispatch-details-modal').classList.remove('hidden');
        }

        function closeDispatchDetailsModal() {
            document.getElementById('dispatch-details-modal').classList.add('hidden');
        }

        function updateStats() {
            const inStock = inventory.filter(i => i.status === 'In').length;
            const dispatched = inventory.filter(i => i.status === 'Out').length;
            
            document.getElementById('stat-total').textContent = inventory.length;
            document.getElementById('stat-in').textContent = inStock;
            document.getElementById('stat-out').textContent = dispatched;

            const categoryData = {};
            const categories = Object.keys(catalog);

            categories.forEach(cat => {
                categoryData[cat] = { total: 0, models: {} };
            });

            inventory.forEach(item => {
                if (item.status === 'In') {
                    const cat = getCategoryFromName(item.name);
                    if (!categoryData[cat]) {
                        categoryData[cat] = { total: 0, models: {} };
                    }
                    categoryData[cat].total++;
                    categoryData[cat].models[item.name] = (categoryData[cat].models[item.name] || 0) + 1;
                }
            });

            const summaryGrid = document.getElementById('category-summary-grid');
            summaryGrid.innerHTML = '';

            Object.keys(categoryData).forEach(cat => {
                const data = categoryData[cat];
                const cardId = 'cat-card-' + cat.replace(/[^a-zA-Z0-9]/g, '_');
                
                let modelsHtml = '';
                const modelNames = Object.keys(data.models);
                if (modelNames.length > 0) {
                    modelNames.forEach(model => {
                        modelsHtml += `<div class="flex justify-between items-center text-xs text-slate-600 py-1 border-b border-slate-100 last:border-0"><span class="truncate pr-2" title="${model}">${model}</span><span class="font-bold bg-slate-200 px-1.5 py-0.5 rounded text-slate-700">${data.models[model]}</span></div>`;
                    });
                } else {
                    modelsHtml = `<div class="text-xs text-slate-400 italic py-1">No items in stock</div>`;
                }

                const div = document.createElement('div');
                div.className = 'bg-slate-50 p-4 rounded-xl border border-slate-200 shadow-sm flex flex-col justify-between';
                div.innerHTML = `
                    <div>
                        <div class="flex justify-between items-start cursor-pointer" onclick="toggleCategoryDropdown('${cardId}')">
                            <div>
                                <span class="text-xs font-bold text-slate-500 uppercase tracking-wide block" title="${cat}">${cat}</span>
                                <span class="text-2xl font-extrabold text-slate-900 mt-1 block">${data.total} <span class="text-xs font-medium text-slate-400">in stock</span></span>
                            </div>
                            <button type="button" class="p-1.5 text-slate-400 hover:text-slate-700 transition"><i id="${cardId}-icon" class="ph ph-caret-down text-base"></i></button>
                        </div>
                    </div>
                    <div id="${cardId}" class="hidden mt-3 pt-3 border-t border-slate-200 space-y-1 max-h-36 overflow-y-auto">
                        ${modelsHtml}
                    </div>
                `;
                summaryGrid.appendChild(div);
            });
        }

        function toggleCategoryDropdown(cardId) {
            const dropdown = document.getElementById(cardId);
            const icon = document.getElementById(cardId + '-icon');
            if (dropdown.classList.contains('hidden')) {
                dropdown.classList.remove('hidden');
                icon.className = 'ph ph-caret-up text-base';
            } else {
                dropdown.classList.add('hidden');
                icon.className = 'ph ph-caret-down text-base';
            }
        }

        function updateMaterialOutSelect() {
            const select = document.getElementById('out-item');
            select.innerHTML = '<option value="" disabled selected>Select equipment...</option>';
            
            const availableItems = inventory.filter(i => i.status === 'In');
            
            if(availableItems.length === 0) {
                select.innerHTML = '<option value="" disabled selected>No equipment available in stock</option>';
                return;
            }

            availableItems.forEach(item => {
                const opt = document.createElement('option');
                opt.value = item.id;
                opt.textContent = `${item.name} (S/N: ${item.sn} | From: ${item.source || 'N/A'})`;
                select.appendChild(opt);
            });
        }

        function formatDate(dateString) {
            if(!dateString) return '';
            const d = new Date(dateString);
            return d.toLocaleDateString('en-GB');
        }

        function openCsvModal() {
            document.getElementById('csv-file-input').value = '';
            document.getElementById('csv-import-modal').classList.remove('hidden');
        }

        function closeCsvModal() {
            document.getElementById('csv-import-modal').classList.add('hidden');
        }

        function downloadCsvTemplate() {
            const csvContent = "FOIF - A90,FOIF - A66,JAVAD - TRIUMPH-3NR\nSN-A90-01,SN-A66-01,SN-33001\nSN-A90-02,,SN-33002";
            const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'janak_horizontal_import_template.csv';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
        }

        function processCsvUpload() {
            const fileInput = document.getElementById('csv-file-input');
            if (!fileInput.files || fileInput.files.length === 0) {
                showToast("Please select a CSV file first.", "error");
                return;
            }

            const file = fileInput.files[0];
            const reader = new FileReader();

            reader.onload = function(e) {
                parseHorizontalCsv(e.target.result);
            };

            reader.readAsText(file);
        }

        function parseHorizontalCsv(text) {
            const lines = text.split(/\r\n|\n/);
            if (lines.length < 2) {
                showToast("CSV file is empty or invalid.", "error");
                return;
            }

            const headers = lines[0].split(',').map(h => h.trim().replace(/^["'](.*)["']$/, '$1'));
            
            let addedCount = 0;
            let skippedCount = 0;
            const todayStr = new Date().toISOString().split('T')[0];

            for (let i = 1; i < lines.length; i++) {
                const line = lines[i].trim();
                if (!line) continue;

                const cols = line.split(',').map(c => c.trim().replace(/^["'](.*)["']$/, '$1'));

                for (let j = 0; j < cols.length; j++) {
                    const sn = (cols[j] || '').toUpperCase();
                    if (!sn) continue;

                    const modelName = headers[j] ? headers[j] : 'Unknown Model';

                    const existing = inventory.find(item => item.sn === sn && item.status === 'In');
                    if (existing) {
                        skippedCount++;
                        continue;
                    }

                    const newItem = {
                        id: 'INV-' + Date.now().toString() + '-' + Math.floor(Math.random() * 10000),
                        name: modelName,
                        sn: sn,
                        source: 'CSV Import',
                        accessories: 'Standard',
                        dateIn: todayStr,
                        status: 'In',
                        dispatchMode: 'Direct',
                        courierName: null,
                        awb: null,
                        company: null,
                        challan: null,
                        phone: null,
                        dispatchTo: null,
                        remarks: null,
                        approvalNotes: null,
                        verificationPhoto: null,
                        dateOut: null,
                        timestamp: Date.now() + addedCount
                    };

                    inventory.push(newItem);
                    addedCount++;
                }
            }

            saveData();
            renderTable();
            updateStats();
            updateMaterialOutSelect();
            checkLowStockLevels();
            closeCsvModal();

            showToast(`Imported ${addedCount} items & synced to Drive!`);
        }

        document.getElementById('form-material-in').addEventListener('submit', function(e) {
            e.preventDefault();
            
            const name = document.getElementById('in-name').value.trim();
            const sn = document.getElementById('in-sn').value.trim().toUpperCase();
            const source = document.getElementById('in-source').value.trim();
            const accessories = document.getElementById('in-accessories').value.trim() || 'None';
            const dateIn = document.getElementById('in-date').value;

            const exists = inventory.find(i => i.sn === sn && i.status === 'In');
            if(exists) {
                showToast(`S/N ${sn} is already currently In Stock!`, "error");
                return;
            }

            const newItem = {
                id: 'INV-' + Date.now().toString(),
                name: name,
                sn: sn,
                source: source,
                accessories: accessories,
                dateIn: dateIn,
                status: 'In',
                dispatchMode: 'Direct',
                courierName: null,
                awb: null,
                company: null,
                challan: null,
                phone: null,
                dispatchTo: null,
                remarks: null,
                approvalNotes: null,
                verificationPhoto: null,
                dateOut: null,
                timestamp: Date.now()
            };

            inventory.push(newItem);
            saveData();
            checkLowStockLevels();
            
            showToast("Equipment added & synced to Drive!");
            this.reset();
            document.getElementById('in-date').valueAsDate = new Date();
            switchTab('dashboard');
        });

        document.getElementById('form-material-out').addEventListener('submit', function(e) {
            e.preventDefault();

            const itemId = document.getElementById('out-item').value;
            const dispatchMode = document.querySelector('input[name="dispatch-mode"]:checked').value;
            const courierName = dispatchMode === 'Courier' ? document.getElementById('out-courier-name').value.trim() : null;
            const awb = dispatchMode === 'Courier' ? (document.getElementById('out-awb').value.trim().toUpperCase() || null) : null;
            const company = document.getElementById('out-company').value.trim();
            const challan = document.getElementById('out-challan').value.trim() || null;
            const phone = document.getElementById('out-phone').value.trim();
            const dispatchTo = document.getElementById('out-to').value.trim();
            const remarks = document.getElementById('out-remarks').value.trim();
            const dateOut = new Date().toISOString();

            const itemIndex = inventory.findIndex(i => i.id === itemId);
            if(itemIndex === -1) {
                showToast("Item not found!", "error");
                return;
            }

            inventory[itemIndex].status = 'Pending Dispatch';
            inventory[itemIndex].dispatchMode = dispatchMode;
            inventory[itemIndex].courierName = courierName;
            inventory[itemIndex].awb = awb;
            inventory[itemIndex].company = company;
            inventory[itemIndex].challan = challan;
            inventory[itemIndex].phone = phone;
            inventory[itemIndex].dispatchTo = dispatchTo;
            inventory[itemIndex].remarks = remarks;
            inventory[itemIndex].dateOut = dateOut;

            saveData();
            showToast("Dispatch requested & synced!");
            this.reset();
            document.querySelector('input[name="dispatch-mode"][value="Direct"]').checked = true;
            toggleDispatchMode();
            updateMaterialOutSelect();
            updateStats();
        });

        function deleteItem(id) {
            inventory = inventory.filter(i => i.id !== id);
            saveData();
            renderTable();
            updateStats();
            checkLowStockLevels();
            showToast("Item deleted.");
        }

        function returnToStock(id) {
            const itemIndex = inventory.findIndex(i => i.id === id);
            if(itemIndex > -1) {
                inventory[itemIndex].status = 'In';
                inventory[itemIndex].dispatchMode = 'Direct';
                inventory[itemIndex].courierName = null;
                inventory[itemIndex].awb = null;
                inventory[itemIndex].company = null;
                inventory[itemIndex].challan = null;
                inventory[itemIndex].phone = null;
                inventory[itemIndex].dispatchTo = null;
                inventory[itemIndex].remarks = null;
                inventory[itemIndex].approvalNotes = null;
                inventory[itemIndex].verificationPhoto = null;
                inventory[itemIndex].dateOut = null;
                inventory[itemIndex].dateIn = new Date().toISOString().split('T')[0];
                
                saveData();
                renderTable();
                updateStats();
                checkLowStockLevels();
                showToast("Item returned to stock & synced.");
            }
        }

        function openEditModal(id) {
            const item = inventory.find(i => i.id === id);
            if(!item) return;

            populateModelDropdowns();
            document.getElementById('edit-id').value = item.id;
            
            const editSelect = document.getElementById('edit-name');
            let existsInSelect = Array.from(editSelect.options).some(opt => opt.value === item.name);
            if (!existsInSelect) {
                const tempOpt = document.createElement('option');
                tempOpt.value = item.name;
                tempOpt.textContent = item.name;
                editSelect.appendChild(tempOpt);
            }
            editSelect.value = item.name;

            document.getElementById('edit-source').value = item.source || '';
            document.getElementById('edit-sn').value = item.sn;
            document.getElementById('edit-accessories').value = item.accessories || '';

            document.getElementById('edit-modal').classList.remove('hidden');
        }

        function closeEditModal() {
            document.getElementById('edit-modal').classList.add('hidden');
        }

        document.getElementById('form-edit').addEventListener('submit', function(e) {
            e.preventDefault();
            const id = document.getElementById('edit-id').value;
            const newName = document.getElementById('edit-name').value.trim();
            const newSource = document.getElementById('edit-source').value.trim();
            const newSn = document.getElementById('edit-sn').value.trim().toUpperCase();
            const newAccessories = document.getElementById('edit-accessories').value.trim();

            const itemIndex = inventory.findIndex(i => i.id === id);
            if(itemIndex > -1) {
                inventory[itemIndex].name = newName;
                inventory[itemIndex].source = newSource;
                inventory[itemIndex].sn = newSn;
                inventory[itemIndex].accessories = newAccessories;
                saveData();
                renderTable();
                updateStats();
                updateMaterialOutSelect();
                closeEditModal();
                showToast("Item updated & synced.");
            }
        });

        function renderApprovals() {
            const tbody = document.getElementById('approvals-table-body');
            tbody.innerHTML = '';
            
            const pendingItems = inventory.filter(i => i.status === 'Pending Dispatch');
            if(pendingItems.length === 0) {
                tbody.innerHTML = `<tr><td colspan="3" class="px-6 py-12 text-center text-slate-500">No pending dispatches.</td></tr>`;
                return;
            }

            pendingItems.forEach(item => {
                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td class="px-6 py-4">
                        <div class="font-semibold text-slate-900">${item.name}</div>
                        <div class="text-sm font-mono text-slate-500">S/N: ${item.sn} (From: ${item.source || 'N/A'})</div>
                    </td>
                    <td class="px-6 py-4">
                        <div class="text-sm font-bold text-slate-800">${item.company || 'N/A'} <span class="font-normal text-xs text-slate-500">[Challan: ${item.challan || 'N/A'}]</span></div>
                        <div class="text-xs text-blue-600 font-semibold mt-0.5">Mode: ${item.dispatchMode}${item.dispatchMode === 'Courier' ? ' ('+item.courierName+')' : ''}</div>
                    </td>
                    <td class="px-6 py-4 text-right">
                        <button type="button" onclick="openApprovalDetailsModal('${item.id}')" class="px-3 py-1.5 bg-emerald-100 text-emerald-700 hover:bg-emerald-200 rounded font-semibold text-sm mr-2 transition">Approve</button>
                        <button type="button" onclick="rejectDispatch('${item.id}')" class="px-3 py-1.5 bg-rose-100 text-rose-700 hover:bg-rose-200 rounded font-semibold text-sm transition">Reject</button>
                    </td>
                `;
                tbody.appendChild(tr);
            });
        }

        function rejectDispatch(id) {
            const itemIndex = inventory.findIndex(i => i.id === id);
            if(itemIndex > -1) {
                inventory[itemIndex].status = 'In';
                inventory[itemIndex].dispatchMode = 'Direct';
                inventory[itemIndex].courierName = null;
                inventory[itemIndex].awb = null;
                inventory[itemIndex].company = null;
                inventory[itemIndex].challan = null;
                inventory[itemIndex].phone = null;
                inventory[itemIndex].dispatchTo = null;
                inventory[itemIndex].remarks = null;
                inventory[itemIndex].approvalNotes = null;
                inventory[itemIndex].verificationPhoto = null;
                inventory[itemIndex].dateOut = null;
                saveData();
                renderApprovals();
                updateStats();
                checkLowStockLevels();
                showToast("Dispatch rejected. Item is back in stock.", "error");
            }
        }

        function openApprovalDetailsModal(id) {
            document.getElementById('approval-item-id').value = id;
            document.getElementById('approval-notes').value = '';
            document.getElementById('approval-details-modal').classList.remove('hidden');
        }

        function closeApprovalDetailsModal() {
            document.getElementById('approval-details-modal').classList.add('hidden');
        }

        async function proceedToCameraCapture() {
            const notes = document.getElementById('approval-notes').value.trim();
            if(!notes) {
                showToast("Please enter verification notes/proof details.", "error");
                return;
            }
            closeApprovalDetailsModal();
            document.getElementById('camera-modal').classList.remove('hidden');
            
            try {
                const stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: 'environment' } });
                streamRef = stream;
                document.getElementById('camera-stream').srcObject = stream;
            } catch(err) {
                showToast("Camera unavailable, approving without photo.", "info");
                skipPhotoAndApprove();
            }
        }

        function skipPhotoAndApprove() {
            const id = document.getElementById('approval-item-id').value;
            const approvalNotes = document.getElementById('approval-notes').value.trim();
            closeCameraModal();
            finalizeDispatch(id, null, approvalNotes);
        }

        function closeCameraModal() {
            document.getElementById('camera-modal').classList.add('hidden');
            if(streamRef) {
                streamRef.getTracks().forEach(track => track.stop());
                streamRef = null;
            }
        }

        function capturePhoto() {
            const video = document.getElementById('camera-stream');
            const canvas = document.getElementById('camera-canvas');
            const id = document.getElementById('approval-item-id').value;
            const approvalNotes = document.getElementById('approval-notes').value.trim();
            
            if(!video.srcObject) {
                skipPhotoAndApprove();
                return;
            }

            canvas.width = video.videoWidth;
            canvas.height = video.videoHeight;
            const ctx = canvas.getContext('2d');
            ctx.drawImage(video, 0, 0, canvas.width, canvas.height);
            
            const photoDataUrl = canvas.toDataURL('image/jpeg', 0.5); 
            
            closeCameraModal();
            finalizeDispatch(id, photoDataUrl, approvalNotes);
        }

        function finalizeDispatch(id, photoDataUrl, approvalNotes) {
            const itemIndex = inventory.findIndex(i => i.id === id);
            if(itemIndex > -1) {
                inventory[itemIndex].status = 'Out';
                inventory[itemIndex].verificationPhoto = photoDataUrl;
                inventory[itemIndex].approvalNotes = approvalNotes;
                saveData();
                
                renderApprovals();
                updateStats();
                checkLowStockLevels();
                showToast("Dispatch Approved & Synced! Gatepass Generated.");
                
                generateGatepass(inventory[itemIndex]);
            }
        }

        function openTrackModal(id) {
            const item = inventory.find(i => i.id === id);
            if(!item) return;

            const content = document.getElementById('track-modal-content');
            
            let html = `
                <div class="mb-4 pb-3 border-b border-slate-200">
                    <h4 class="font-bold text-slate-900">${item.name}</h4>
                    <p class="text-xs text-slate-500 font-mono">Serial Number: ${item.sn}</p>
                </div>
                <div class="space-y-6 relative before:absolute before:left-3.5 before:top-2 before:bottom-2 before:w-0.5 before:bg-slate-200 pl-8">
                    <!-- Step 1: Material In -->
                    <div class="relative">
                        <div class="absolute -left-8 top-0.5 w-7 h-7 rounded-full bg-emerald-100 text-emerald-700 flex items-center justify-center font-bold text-xs border-2 border-white shadow">1</div>
                        <div>
                            <p class="text-xs font-bold text-emerald-700 uppercase tracking-wide">Received In Stock</p>
                            <p class="text-sm font-semibold text-slate-800 mt-0.5">Source: ${item.source || 'N/A'}</p>
                            <p class="text-xs text-slate-500">Accessories: ${item.accessories || 'None'}</p>
                            <p class="text-xs text-slate-500">Date: ${formatDate(item.dateIn)}</p>
                        </div>
                    </div>
            `;

            if (item.status === 'Out' || item.status === 'Pending Dispatch') {
                const isOut = item.status === 'Out';
                html += `
                    <!-- Step 2: Dispatch Request / Approval -->
                    <div class="relative">
                        <div class="absolute -left-8 top-0.5 w-7 h-7 rounded-full ${isOut ? 'bg-blue-100 text-blue-700' : 'bg-purple-100 text-purple-700'} flex items-center justify-center font-bold text-xs border-2 border-white shadow">2</div>
                        <div>
                            <p class="text-xs font-bold ${isOut ? 'text-blue-700' : 'text-purple-700'} uppercase tracking-wide">${isOut ? 'Dispatched' : 'Pending Dispatch Request'}</p>
                            <p class="text-sm font-semibold text-slate-800 mt-0.5">Client: ${item.company || 'N/A'} [Challan: ${item.challan || 'N/A'}]</p>
                            <p class="text-xs text-slate-500">Location: ${item.dispatchTo || 'N/A'} | Phone: ${item.phone || 'N/A'}</p>
                            <p class="text-xs text-slate-500">Mode: ${item.dispatchMode || 'Direct'}${item.dispatchMode === 'Courier' ? ' ('+item.courierName+' | AWB: '+(item.awb || 'N/A')+')' : ''}</p>
                        </div>
                    </div>
                `;

                if (isOut && item.approvalNotes) {
                    html += `
                        <!-- Step 3: Verification & Proof -->
                        <div class="relative">
                            <div class="absolute -left-8 top-0.5 w-7 h-7 rounded-full bg-emerald-100 text-emerald-700 flex items-center justify-center font-bold text-xs border-2 border-white shadow">3</div>
                            <div>
                                <p class="text-xs font-bold text-emerald-700 uppercase tracking-wide">Admin Verified & Approved</p>
                                <p class="text-xs text-slate-700 mt-1 italic bg-slate-50 p-2 rounded border border-slate-200">"${item.approvalNotes}"</p>
                                ${item.verificationPhoto ? `<div class="mt-2"><img src="${item.verificationPhoto}" class="w-32 h-24 object-cover rounded-lg border shadow-sm"></div>` : ''}
                            </div>
                        </div>
                    `;
                }
            } else {
                html += `
                    <div class="relative">
                        <div class="absolute -left-8 top-0.5 w-7 h-7 rounded-full bg-slate-200 text-slate-600 flex items-center justify-center font-bold text-xs border-2 border-white shadow">2</div>
                        <div>
                            <p class="text-xs font-bold text-slate-600 uppercase tracking-wide">Current Status</p>
                            <p class="text-sm font-medium text-slate-700 mt-0.5">Available in stock for assignment.</p>
                        </div>
                    </div>
                `;
            }

            html += `</div>`;
            content.innerHTML = html;
            document.getElementById('track-modal').classList.remove('hidden');
        }

        function closeTrackModal() {
            document.getElementById('track-modal').classList.add('hidden');
        }

        function generateGatepass(item) {
            const gpNo = "GP-" + Math.floor(10000 + Math.random() * 90000);
            const dateStr = formatDate(new Date().toISOString());

            document.getElementById('print-gp-no').textContent = gpNo;
            document.getElementById('print-date').textContent = dateStr;
            document.getElementById('print-challan').textContent = item.challan || 'N/A';
            document.getElementById('print-company').textContent = item.company || 'N/A';
            document.getElementById('print-to').textContent = "Site / Location: " + (item.dispatchTo || 'N/A');
            document.getElementById('print-phone').textContent = item.phone || 'N/A';
            
            const courierInfoEl = document.getElementById('print-courier-info');
            if (item.dispatchMode === 'Courier') {
                courierInfoEl.textContent = `Courier Partner: ${item.courierName} | AWB No: ${item.awb || 'N/A'}`;
                courierInfoEl.classList.remove('hidden');
            } else {
                courierInfoEl.classList.add('hidden');
            }

            document.getElementById('print-item-name').textContent = item.name + " (Received From: " + (item.source || 'N/A') + ")";
            document.getElementById('print-item-sn').textContent = item.sn;
            document.getElementById('print-remarks').textContent = item.remarks || "None";

            const printSource = document.getElementById('gatepass-print-view');
            const previewTarget = document.getElementById('preview-document');
            previewTarget.innerHTML = printSource.innerHTML;

            document.getElementById('gatepass-preview-modal').classList.remove('hidden');
        }

        function closeGatepassPreview() {
            document.getElementById('gatepass-preview-modal').classList.add('hidden');
            switchTab('dashboard'); 
        }

        window.addEventListener('DOMContentLoaded', checkAuth);
    </script>
</body>
</html>
```<ElicitationsGroup></ElicitationsGroup>
