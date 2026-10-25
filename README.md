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
        html, body {
            font-family: 'Inter', sans-serif;
            background-color: #f8fafc;
            margin: 0 !important;
            padding: 0 !important;
            width: 100vw !important;
            height: 100vh !important;
            max-width: 100vw !important;
            overflow: hidden !important;
        }

        /* Escape any GitHub Pages / Jekyll markdown wrapper if rendered inside README */
        .container-lg, .markdown-body, .wrapper, .page-content {
            max-width: 100vw !important;
            width: 100vw !important;
            margin: 0 !important;
            padding: 0 !important;
        }
        .markdown-body > h1:first-child {
            display: none !important;
        }
        
        /* Custom Scrollbar for Main Content Area */
        ::-webkit-scrollbar { width: 10px; height: 10px; }
        ::-webkit-scrollbar-track { background: #e2e8f0; }
        ::-webkit-scrollbar-thumb { background: #64748b; border-radius: 5px; border: 2px solid #e2e8f0; }
        ::-webkit-scrollbar-thumb:hover { background: #334155; }

        /* Sidebar Scrollbar */
        .sidebar-scroll::-webkit-scrollbar { width: 6px; height: 6px; }
        .sidebar-scroll::-webkit-scrollbar-track { background: #0f172a; }
        .sidebar-scroll::-webkit-scrollbar-thumb { background: #334155; border-radius: 4px; }
        .sidebar-scroll::-webkit-scrollbar-thumb:hover { background: #475569; }

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
<body class="text-slate-800 antialiased h-screen w-screen overflow-hidden flex bg-slate-50">

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
    <div id="app-container" class="hidden fixed inset-0 z-40 w-screen h-screen flex-col md:flex-row bg-slate-50 overflow-hidden">
        
        <!-- Sidebar -->
        <aside class="w-full md:w-64 md:h-full bg-slate-900 text-white flex flex-col shadow-xl z-20 shrink-0 overflow-hidden">
            <div class="p-6 border-b border-slate-800 flex flex-col justify-center">
                <h1 class="font-black text-xl leading-tight text-red-600 uppercase tracking-wide whitespace-nowrap">Janak Positioning</h1>
                <p class="text-slate-400 font-medium text-xs mt-1 uppercase tracking-widest">Surveying Systems</p>
            </div>
            
            <nav class="flex-1 py-4 px-3 flex md:flex-col gap-2 overflow-x-auto md:overflow-y-auto sidebar-scroll">
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

            <!-- Scrollable Viewport Area -->
            <div id="main-scroll-area" class="flex-1 h-full overflow-y-scroll overflow-x-hidden p-4 md:p-6 lg:p-8 relative">
                
                <!-- Tab: Dashboard -->
                <div id="tab-dashboard" class="tab-content w-full block space-y-6 pb-16">
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
                <div id="tab-material-in" class="tab-content w-full max-w-4xl mx-auto hidden pb-16">
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
                                <select id="in-name" required class="w-full px-4 py-3 border border-slate-300 rounded-xl focus:ring-
