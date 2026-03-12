from django.urls import path
from . import views

urlpatterns = [
    # Auth
    path('login/', views.login_page, name='login'),
    path('logout/', views.logout_view, name='logout'),
    path('register/', views.register_page, name='register'),

    # Main
    path('', views.index, name='index'),
    path('sheet/<int:sheet_id>/', views.sheet_view, name='sheet_view'),

    # API - sheet data
    path('api/sheet/<int:sheet_id>/data/', views.get_sheet_data, name='sheet_data'),
    path('api/sheet/<int:sheet_id>/update/', views.update_cell, name='update_cell'),
    path('api/sheet/<int:sheet_id>/comment/', views.update_comment, name='update_comment'),
    path('api/sheet/<int:sheet_id>/search/', views.search_cells, name='search_cells'),
    path('api/sheet/<int:sheet_id>/cell/<int:row>/<int:col>/history/', views.cell_history, name='cell_history'),
    path('api/sheet/<int:sheet_id>/export/', views.export_excel, name='export_excel'),

    # Snapshots
    path('api/sheet/<int:sheet_id>/snapshots/', views.list_snapshots, name='list_snapshots'),
    path('api/sheet/<int:sheet_id>/snapshots/create/', views.create_snapshot, name='create_snapshot'),
    path('api/sheet/<int:sheet_id>/snapshots/<int:snapshot_id>/restore/', views.restore_snapshot, name='restore_snapshot'),

    # Compare
    path('compare/', views.compare_sheets, name='compare_sheets'),
    path('api/compare/', views.compare_api, name='compare_api'),

    # Import / merge
    path('api/import/', views.import_excel_view, name='import_excel'),
    path('api/excel-sheets/', views.excel_sheets_list, name='excel_sheets_list'),
    path('api/merge/', views.merge_files, name='merge_files'),
    path('api/apply-data/', views.apply_data_view, name='apply_data'),

    # Pages
    path('merge/', views.merge_page, name='merge_page'),
    path('apply-data/', views.apply_data_page, name='apply_data_page'),
    path('seed/', views.seed_demo, name='seed_demo'),
]
