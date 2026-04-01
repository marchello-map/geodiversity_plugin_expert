# -*- coding: utf-8 -*-
"""
/***************************************************************************
 Geodiversity Calculator v2.1
                                 A QGIS plugin
 The world's most advanced, robust geodiversity assessment plugin.
 Handles national-scale datasets with bulletproof error handling.
 Based on the proven methodology by Márton Pál, enhanced for v2.1.
                              -------------------
        begin                : 2026-01-26
        copyright            : (C) 2026 v2.1 Team (enhanced from Márton Pál by Márton Pál and Emmanuel Owusu-Acheampong)
        email                : pal.marton@inf.elte.hu & emmaoacheamp@student.elte.hu
 ***************************************************************************/
"""
from qgis.PyQt.QtCore import QSettings, QTranslator, QCoreApplication, QVariant, QDateTime
from qgis.PyQt.QtGui import QIcon
from qgis.PyQt.QtWidgets import QAction, QFileDialog, QDialogButtonBox, QProgressBar
from qgis.core import (
    Qgis, QgsVectorLayer, QgsProject, QgsField, QgsProcessingFeedback,
    QgsRasterLayer, QgsMessageLog, QgsVectorLayerJoinInfo, QgsSpatialIndex,
    QgsExpression, QgsExpressionContext, QgsExpressionContextScope, edit,
    QgsStyle, QgsGraduatedSymbolRenderer
)
from .geodiversity_calculator_dialog import GeodiversityCalculatorDialog
import os
import os.path
import processing
import math
import traceback

class GeodiversityCalculator:
    """Geodiversity calculator"""
    
    def __init__(self, iface):
        self.iface = iface
        self.plugin_dir = os.path.dirname(__file__)
        self.actions = []
        self.first_start = True
        self.dlg = None
        
    def tr(self, message):
        return QCoreApplication.translate('GeodiversityCalculator', message)
    
    def log(self, message, level=Qgis.Info):
        """Enhanced logging"""
        QgsMessageLog.logMessage(message, 'Geodiversity Calculator', level)

    def _layer_base_name(self, path: str, fallback: str = "layer") -> str:
        """Return file base name without extension from a path."""
        try:
            if not path:
                return fallback
            return os.path.splitext(os.path.basename(path))[0]
        except Exception:
            return fallback

    def _output_gpkg_path(self, working_dir: str, source_path: str, suffix: str = "_grid", fallback: str = "layer") -> str:
        """Build an output gpkg path based on the input layer name."""
        base_name = self._layer_base_name(source_path, fallback)
        return os.path.join(working_dir, f"{base_name}{suffix}.gpkg")

    def _set_progress(self, progressbar, value: int):
        """Update progress bar and let the UI breathe."""
        try:
            if progressbar:
                progressbar.setValue(int(value))
                QCoreApplication.processEvents()
        except Exception:
            pass

    def _remove_field_if_exists(self, layer: QgsVectorLayer, field_name: str):
        """Physically remove a field from a layer (datasource) if present.

        Note: this modifies the source layer on disk (shp/gpkg).
        """
        try:
            if layer is None or not layer.isValid():
                return
            idx = layer.fields().lookupField(field_name)
            if idx == -1:
                return
            layer.dataProvider().deleteAttributes([idx])
            layer.updateFields()
        except Exception:
            return

    def _encode_unique_values(self, layer: QgsVectorLayer, src_field: str, out_field: str = "r_value"):
        """Add/overwrite an integer category field based on unique values in src_field.

        Returns a dict mapping original values -> integer codes (starting at 1).
        """
        if layer is None or not layer.isValid():
            raise Exception("Layer invalid for encoding unique values.")
        if not src_field:
            raise Exception("Source field not provided for encoding unique values.")
        if layer.fields().lookupField(out_field) == -1:
            layer.dataProvider().addAttributes([QgsField(out_field, QVariant.Int)])
            layer.updateFields()

        mapping = {}
        next_code = 1
        out_idx = layer.fields().lookupField(out_field)

        with edit(layer):
            for f in layer.getFeatures():
                key = f[src_field]
                if key not in mapping:
                    mapping[key] = next_code
                    next_code += 1
                f.setAttribute(out_idx, mapping[key])
                layer.updateFeature(f)
        return mapping

    def _vector_touch_variety(self, grid_path: str, poly_layer: QgsVectorLayer, out_path: str,
                             out_field: str, value_field: str = "r_value"):
        """Backward-compatible wrapper for geometry-intersection variety."""
        return self._vector_touch_variety_any_geometry(
            grid_path=grid_path,
            input_layer=poly_layer,
            out_path=out_path,
            out_field=out_field,
            value_field=value_field
        )

    def _vector_touch_variety_any_geometry(self, grid_path: str, input_layer: QgsVectorLayer, out_path: str,
                                           out_field: str, value_field: str = "r_value"):
        """Compute distinct-category variety per grid cell for any vector geometry type.

        Variety is the count of DISTINCT values (value_field) among features whose geometries
        INTERSECT the grid cell geometry. This works for polygons, lines and points.
        """
        if not grid_path or not os.path.exists(grid_path):
            raise Exception("Grid path not found for vector variety computation.")
        if input_layer is None or not input_layer.isValid():
            raise Exception("Input layer invalid for vector variety computation.")

        processing.run("native:savefeatures", {
            'INPUT': grid_path,
            'OUTPUT': out_path
        })

        out_grid = QgsVectorLayer(out_path, os.path.basename(out_path), "ogr")
        if not out_grid.isValid():
            raise Exception("Failed to create output grid for vector variety computation.")

        prov = out_grid.dataProvider()
        if out_grid.fields().lookupField(out_field) == -1:
            prov.addAttributes([QgsField(out_field, QVariant.Int)])
            out_grid.updateFields()
        out_idx = out_grid.fields().lookupField(out_field)

        sidx = QgsSpatialIndex(input_layer.getFeatures())

        with edit(out_grid):
            for cell in out_grid.getFeatures():
                g = cell.geometry()
                if g is None or g.isEmpty():
                    prov.changeAttributeValues({cell.id(): {out_idx: 0}})
                    continue

                cand_ids = sidx.intersects(g.boundingBox())
                vals = set()
                for fid in cand_ids:
                    in_feat = input_layer.getFeature(fid)
                    ig = in_feat.geometry()
                    if ig is None or ig.isEmpty():
                        continue
                    if ig.intersects(g):
                        vals.add(in_feat[value_field])
                prov.changeAttributeValues({cell.id(): {out_idx: int(len(vals))}})

        out_grid.updateExtents()
        return out_grid

    def _geomorphology_vector_variety(self, grid_path: str, out_path: str, geom_inputs: list):
        """Compute combined geomorphology variety from optional line/polygon/point layers.

        geom_inputs items: (path, class_field, layer_name)
        Distinct classes are counted across all provided layers.
        """
        prepared = []
        code_offset = 0

        for path, cls_field, layer_name in geom_inputs:
            if not path:
                continue
            layer = QgsVectorLayer(path, layer_name, "ogr")
            if not layer.isValid():
                raise Exception(f"Invalid geomorphology layer: {layer_name}")
            if not cls_field:
                raise Exception(f"Classification field is required for {layer_name}.")

            mapping = self._encode_unique_values(layer, cls_field, out_field='r_value')
            offset_map = {orig_val: (code + code_offset) for orig_val, code in mapping.items()}
            code_offset += len(mapping)
            prepared.append({
                'layer': layer,
                'class_field': cls_field,
                'code_map': offset_map,
                'index': QgsSpatialIndex(layer.getFeatures())
            })

        if not prepared:
            return None

        processing.run("native:savefeatures", {
            'INPUT': grid_path,
            'OUTPUT': out_path
        })
        out_grid = QgsVectorLayer(out_path, os.path.basename(out_path), "ogr")
        if not out_grid.isValid():
            raise Exception("Failed to create output grid for geomorphology variety computation.")

        prov = out_grid.dataProvider()
        if out_grid.fields().lookupField('_geom_variety') == -1:
            prov.addAttributes([QgsField('_geom_variety', QVariant.Int)])
            out_grid.updateFields()
        out_idx = out_grid.fields().lookupField('_geom_variety')

        with edit(out_grid):
            for cell in out_grid.getFeatures():
                g = cell.geometry()
                vals = set()
                if g is not None and not g.isEmpty():
                    for item in prepared:
                        layer = item['layer']
                        sidx = item['index']
                        cand_ids = sidx.intersects(g.boundingBox())
                        for fid in cand_ids:
                            feat = layer.getFeature(fid)
                            fg = feat.geometry()
                            if fg is None or fg.isEmpty() or not fg.intersects(g):
                                continue
                            orig_val = feat[item['class_field']]
                            vals.add(item['code_map'].get(orig_val))
                prov.changeAttributeValues({cell.id(): {out_idx: int(len([v for v in vals if v is not None]))}})

        for item in prepared:
            self._remove_field_if_exists(item['layer'], 'r_value')

        out_grid.updateExtents()
        return out_grid

    def _suggest_grid_spacing_from_boundary(self, boundary_path: str):
        """Auto-suggest grid spacing based on boundary extent area (km²)."""
        try:
            if hasattr(self, "_spacing_autofill") and not self._spacing_autofill:
                return

            boundary_path = (boundary_path or "").strip()
            if not boundary_path or not os.path.exists(boundary_path):
                return

            layer = QgsVectorLayer(boundary_path, "boundary", "ogr")
            if not layer.isValid():
                return
            area_km2 = layer.extent().width() * layer.extent().height() / 1_000_000.0
            area = area_km2
            if area <= 0:
                return

            if area > 5_000_000:
                suggested = 50000
            elif 1_000_000 < area <= 5_000_000:
                suggested = 20000
            elif 100_000 < area <= 1_000_000:
                suggested = 10000
            elif 50_000 < area <= 100_000:
                suggested = 7500
            elif 10_000 < area <= 50_000:
                suggested = 5000
            elif 2_500 < area <= 10_000:
                suggested = 2500
            else:
                suggested = 1000

            self.dlg.lineEdit_3.setText(str(suggested))
            self.dlg.lineEdit_4.setText(str(suggested))
        except Exception:
            return

    def _add_normalized_fields(self, grid):
        """Create normalized subindex fields (0-1) and their sum (N_sum)."""
        src = {
            "N_geol": "J_geol_variety",
            "N_pedo": "J_pedo_variety",
            "N_geom": "J_geom_variety",
            "N_miner": "J_mineral_idx",
            "N_foss": "J_fossil_idx",
        }

        max_vals = {k: 0.0 for k in src.keys()}
        max_hydro = 0.0

        for f in grid.getFeatures():
            for out_name, in_name in src.items():
                v = f[in_name]
                try:
                    v = float(v) if v is not None else 0.0
                except Exception:
                    v = 0.0
                if v > max_vals[out_name]:
                    max_vals[out_name] = v

            stra = f["J_stra_max"]
            lakes = f["_lakes"]
            try:
                stra = float(stra) if stra is not None else 0.0
            except Exception:
                stra = 0.0
            try:
                lakes = float(lakes) if lakes is not None else 0.0
            except Exception:
                lakes = 0.0
            hydro = stra + lakes
            if hydro > max_hydro:
                max_hydro = hydro

        for k in max_vals:
            if max_vals[k] <= 0:
                max_vals[k] = 0.0
        if max_hydro <= 0:
            max_hydro = 0.0

        prov = grid.dataProvider()
        to_add = []
        for fn in list(src.keys()) + ["N_hydro", "N_sum"]:
            if grid.fields().lookupField(fn) == -1:
                to_add.append(QgsField(fn, QVariant.Double))
        if to_add:
            prov.addAttributes(to_add)
            grid.updateFields()

        idxs = {fn: grid.fields().lookupField(fn) for fn in list(src.keys()) + ["N_hydro", "N_sum"]}

        with edit(grid):
            for f in grid.getFeatures():
                vals = {}
                n_sum = 0.0

                for out_name, in_name in src.items():
                    raw = f[in_name]
                    try:
                        raw = float(raw) if raw is not None else 0.0
                    except Exception:
                        raw = 0.0
                    mx = max_vals[out_name]
                    n = (raw / mx) if mx and mx > 0 else 0.0
                    vals[idxs[out_name]] = n
                    n_sum += n

                stra = f["J_stra_max"]
                lakes = f["_lakes"]
                try:
                    stra = float(stra) if stra is not None else 0.0
                except Exception:
                    stra = 0.0
                try:
                    lakes = float(lakes) if lakes is not None else 0.0
                except Exception:
                    lakes = 0.0
                hydro_raw = stra + lakes
                n_hydro = (hydro_raw / max_hydro) if max_hydro and max_hydro > 0 else 0.0
                vals[idxs["N_hydro"]] = n_hydro
                n_sum += n_hydro

                vals[idxs["N_sum"]] = n_sum
                grid.dataProvider().changeAttributeValues({f.id(): vals})

    def _apply_output_style(self, layer, field_name: str):
        """Apply an initial graduated style (Reds ramp, Jenks)."""
        try:
            if layer is None or not layer.isValid():
                return
            if layer.fields().lookupField(field_name) == -1:
                return

            style = QgsStyle.defaultStyle()
            ramp = style.colorRamp("Reds") if style else None
            if ramp is None:
                return

            renderer = QgsGraduatedSymbolRenderer()
            renderer.setClassAttribute(field_name)
            renderer.setSourceColorRamp(ramp)

            renderer.updateClasses(layer, QgsGraduatedSymbolRenderer.Jenks, 5)
            renderer.updateColorRamp(ramp)

            layer.setRenderer(renderer)
            layer.triggerRepaint()
        except Exception:
            return
        
    def add_action(self, icon_path, text, callback, parent=None):
        icon = QIcon(icon_path)
        action = QAction(icon, text, parent)
        action.triggered.connect(callback)
        self.iface.addToolBarIcon(action)
        self.iface.addPluginToMenu(self.tr(u'&Geodiversity Calculator v2.1'), action)
        self.actions.append(action)
        return action
    
    def initGui(self):
        icon_path = os.path.join(self.plugin_dir, 'icon.png')
        self.add_action(icon_path, text=self.tr(u'Geodiversity Calculator v2.1'),
                       callback=self.run, parent=self.iface.mainWindow())
    
    def unload(self):
        for action in self.actions:
            self.iface.removePluginMenu(self.tr(u'&Geodiversity Calculator v2.1'), action)
            self.iface.removeToolBarIcon(action)
    
    def select_result_folder(self):
        folder = QFileDialog.getExistingDirectory(self.dlg, "Select result folder")
        if folder:
            self.dlg.lineEdit_16.setText(folder)
    
    def select_boundary(self):
        filename, _ = QFileDialog.getOpenFileName(
            self.dlg, "Select boundary", "", "Vector (*.gpkg *.shp)"
        )
        if filename:
            self.dlg.lineEdit.setText(filename)
            self._suggest_grid_spacing_from_boundary(filename)
    
    def select_geology(self):
        filename, _ = QFileDialog.getOpenFileName(
            self.dlg, "Select geology", "", "Vector (*.gpkg *.shp)"
        )
        if filename:
            self.dlg.lineEdit_5.setText(filename)
    
    def select_pedology(self):
        filename, _ = QFileDialog.getOpenFileName(
            self.dlg, "Select pedology", "", "Vector (*.gpkg *.shp)"
        )
        if filename:
            self.dlg.lineEdit_7.setText(filename)
    
    def select_dem(self):
        filename, _ = QFileDialog.getOpenFileName(
            self.dlg, "Select DEM", "", "Raster (*.tif)"
        )
        if filename:
            self.dlg.lineEdit_9.setText(filename)

    def select_geom_line(self):
        filename, _ = QFileDialog.getOpenFileName(
            self.dlg, "Select geomorphology line layer", "", "Vector (*.gpkg *.shp)"
        )
        if filename:
            self.dlg.lineEdit_geom_line.setText(filename)

    def select_geom_poly(self):
        filename, _ = QFileDialog.getOpenFileName(
            self.dlg, "Select geomorphology polygon layer", "", "Vector (*.gpkg *.shp)"
        )
        if filename:
            self.dlg.lineEdit_geom_poly.setText(filename)

    def select_geom_point(self):
        filename, _ = QFileDialog.getOpenFileName(
            self.dlg, "Select geomorphology point layer", "", "Vector (*.gpkg *.shp)"
        )
        if filename:
            self.dlg.lineEdit_geom_point.setText(filename)

    def select_geomorphon_raster(self):
        filename, _ = QFileDialog.getOpenFileName(
            self.dlg, "Select geomorphon raster", "", "Raster (*.tif *.img *.asc *.sdat)"
        )
        if filename:
            self.dlg.lineEdit_geomorphon_raster.setText(filename)
    
    def select_lakes(self):
        filename, _ = QFileDialog.getOpenFileName(
            self.dlg, "Select lakes/seas", "", "Vector (*.gpkg *.shp)"
        )
        if filename:
            self.dlg.lineEdit_17.setText(filename)
    
    def select_mineral(self):
        filename, _ = QFileDialog.getOpenFileName(
            self.dlg, "Select mineralogy", "", "Vector (*.gpkg *.shp)"
        )
        if filename:
            self.dlg.lineEdit_12.setText(filename)
    
    def select_palaeo(self):
        filename, _ = QFileDialog.getOpenFileName(
            self.dlg, "Select palaeontology", "", "Vector (*.gpkg *.shp)"
        )
        if filename:
            self.dlg.lineEdit_13.setText(filename)
    
    def clear_all(self):
        widgets = [
            self.dlg.lineEdit_16, self.dlg.lineEdit, self.dlg.lineEdit_3, self.dlg.lineEdit_4,
            self.dlg.lineEdit_5, self.dlg.lineEdit_6, self.dlg.lineEdit_7, self.dlg.lineEdit_8,
            self.dlg.lineEdit_9, self.dlg.lineEdit_17, self.dlg.lineEdit_12, self.dlg.lineEdit_10,
            self.dlg.lineEdit_13, self.dlg.lineEdit_11, self.dlg.lineEdit_14,
            getattr(self.dlg, 'lineEdit_geom_line', None), getattr(self.dlg, 'lineEdit_geom_line_field', None),
            getattr(self.dlg, 'lineEdit_geom_poly', None), getattr(self.dlg, 'lineEdit_geom_poly_field', None),
            getattr(self.dlg, 'lineEdit_geom_point', None), getattr(self.dlg, 'lineEdit_geom_point_field', None),
            getattr(self.dlg, 'lineEdit_geomorphon_raster', None)
        ]
        for widget in widgets:
            if widget is not None:
                widget.clear()
    
    def run(self):
        if self.first_start:
            self.first_start = False
            self.dlg = GeodiversityCalculatorDialog()
            def _sync_geom_ui():
                try:
                    if hasattr(self.dlg, 'comboBox_geom_source') and hasattr(self.dlg, 'stackedWidget_geom'):
                        idx = int(self.dlg.comboBox_geom_source.currentIndex())
                        self.dlg.stackedWidget_geom.setCurrentIndex(idx)
                        self.dlg.stackedWidget_geom.setVisible(True)
                except Exception:
                    pass

            try:
                if hasattr(self.dlg, 'comboBox_geom_source') and hasattr(self.dlg, 'stackedWidget_geom'):
                    self.dlg.comboBox_geom_source.currentIndexChanged.connect(lambda _i: _sync_geom_ui())
                    _sync_geom_ui()
                else:
                    self.log('Expert geomorphology widgets not found in UI (comboBox_geom_source/stackedWidget_geom). '
                             'Make sure geodiversity_calculator_dialog_base.ui was overwritten.', Qgis.Warning)
            except Exception:
                pass

            try:
                self.dlg.lineEdit_3.clear()
                self.dlg.lineEdit_4.clear()
            except Exception:
                pass

            self._spacing_autofill = True
            try:
                self.dlg.lineEdit_3.textEdited.connect(lambda _t: setattr(self, "_spacing_autofill", False))
                self.dlg.lineEdit_4.textEdited.connect(lambda _t: setattr(self, "_spacing_autofill", False))
            except Exception:
                pass

            try:
                self.dlg.lineEdit.textChanged.connect(self._suggest_grid_spacing_from_boundary)
            except Exception:
                pass
            self.dlg.pushButton_14.clicked.connect(self.select_result_folder)
            self.dlg.pushButton.clicked.connect(self.select_boundary)
            self.dlg.pushButton_3.clicked.connect(self.select_geology)
            self.dlg.pushButton_5.clicked.connect(self.select_pedology)
            self.dlg.pushButton_7.clicked.connect(self.select_dem)
            self.dlg.pushButton_13.clicked.connect(self.select_lakes)
            self.dlg.pushButton_10.clicked.connect(self.select_mineral)
            self.dlg.pushButton_11.clicked.connect(self.select_palaeo)
            try:
                self.dlg.pushButton_geom_line.clicked.connect(self.select_geom_line)
                self.dlg.pushButton_geom_poly.clicked.connect(self.select_geom_poly)
                self.dlg.pushButton_geom_point.clicked.connect(self.select_geom_point)
                self.dlg.pushButton_geomorphon_raster.clicked.connect(self.select_geomorphon_raster)
            except Exception:
                pass
            self.dlg.button_box.button(QDialogButtonBox.Reset).clicked.connect(self.clear_all)
        
        self.dlg.show()
        if self.dlg.exec_():
            self.execute()
    
    def execute(self):
        """Main execution"""
        progressbar = None
        try:
            self.iface.messageBar().clearWidgets()
            progressbar = QProgressBar()
            self.iface.messageBar().pushWidget(progressbar)
            self._set_progress(progressbar, 0)
            
            working_dir = self.dlg.lineEdit_16.text().strip()
            boundary_path = self.dlg.lineEdit.text().strip()
            grid_name = self.dlg.lineEdit_14.text().strip()
            
            try:
                h_txt = (self.dlg.lineEdit_3.text() or '').strip()
                v_txt = (self.dlg.lineEdit_4.text() or '').strip()
                if not h_txt or not v_txt:
                    raise Exception('Grid spacing values are required (set manually or select a boundary for auto-fill).')
                h_spacing = float(h_txt)
                v_spacing = float(v_txt)
            except ValueError:
                raise Exception("Invalid grid spacing values!")
            
            geology_path = self.dlg.lineEdit_5.text().strip() or None
            geol_field = self.dlg.lineEdit_6.text().strip() or None
            pedology_path = self.dlg.lineEdit_7.text().strip() or None
            pedo_field = self.dlg.lineEdit_8.text().strip() or None
            dem_path = self.dlg.lineEdit_9.text().strip() or None
            lakes_path = self.dlg.lineEdit_17.text().strip() or None
            geom_source_idx = int(self.dlg.comboBox_geom_source.currentIndex()) if hasattr(self.dlg, 'comboBox_geom_source') else 1
            geom_line_path = self.dlg.lineEdit_geom_line.text().strip() if hasattr(self.dlg, 'lineEdit_geom_line') else None
            geom_line_field = self.dlg.lineEdit_geom_line_field.text().strip() if hasattr(self.dlg, 'lineEdit_geom_line_field') else None
            geom_poly_path = self.dlg.lineEdit_geom_poly.text().strip() if hasattr(self.dlg, 'lineEdit_geom_poly') else None
            geom_poly_field = self.dlg.lineEdit_geom_poly_field.text().strip() if hasattr(self.dlg, 'lineEdit_geom_poly_field') else None
            geom_point_path = self.dlg.lineEdit_geom_point.text().strip() if hasattr(self.dlg, 'lineEdit_geom_point') else None
            geom_point_field = self.dlg.lineEdit_geom_point_field.text().strip() if hasattr(self.dlg, 'lineEdit_geom_point_field') else None
            geomorphon_raster_path = self.dlg.lineEdit_geomorphon_raster.text().strip() if hasattr(self.dlg, 'lineEdit_geomorphon_raster') else None
            mineral_path = self.dlg.lineEdit_12.text().strip() or None
            mineral_field = self.dlg.lineEdit_10.text().strip() or None
            palaeo_path = self.dlg.lineEdit_13.text().strip() or None
            palaeo_field = self.dlg.lineEdit_11.text().strip() or None
            
            if not all([boundary_path, working_dir, grid_name]):
                raise Exception("Required fields missing!")
            
            try:
                show_intermediate = bool(getattr(self.dlg, "checkBox_show_sublayers", None) and self.dlg.checkBox_show_sublayers.isChecked())
            except Exception:
                show_intermediate = False

            def _add_layer(layer, intermediate: bool = True):
                try:
                    if layer is None or not layer.isValid():
                        return
                    if (not intermediate) or show_intermediate:
                        QgsProject.instance().addMapLayer(layer)
                except Exception:
                    return
            
            self.log("Starting Geodiversity Calculator v2.1 analysis...", Qgis.Info)
            
            self.log("Creating analysis grid...", Qgis.Info)
            grid0 = working_dir + "/" + grid_name + ".gpkg"
            grid0_temp = None
            hatar0 = QgsVectorLayer(boundary_path, "boundary", "ogr")
            
            if not hatar0.isValid():
                raise Exception("Boundary layer is invalid!")
            
            _add_layer(hatar0, intermediate=True)
            
            boundary_area_km2 = hatar0.extent().width() * hatar0.extent().height() / 1_000_000
            self.log(f"Boundary extent area: {boundary_area_km2:.0f} km²", Qgis.Info)
            
            area = boundary_area_km2
            if area > 5_000_000:
                suggested_spacing = 50000
                self.log(f"HUGE dataset detected ({area:,.0f} km²). Consider using {suggested_spacing} m grid spacing.", Qgis.Warning)
            elif 1_000_000 < area <= 5_000_000:
                suggested_spacing = 20000
                self.log(f"Large dataset detected ({area:,.0f} km²). Consider using {suggested_spacing} m grid spacing.", Qgis.Warning)
            elif 100_000 < area <= 1_000_000:
                suggested_spacing = 10000
                self.log(f"Medium dataset detected ({area:,.0f} km²). Consider using {suggested_spacing} m grid spacing.", Qgis.Info)
            elif 50_000 < area <= 100_000:
                suggested_spacing = 5000
                self.log(f"Moderate dataset detected ({area:,.0f} km²). Consider using {suggested_spacing} m grid spacing.", Qgis.Info)
            elif 20_000 < area <= 50_000:
                suggested_spacing = 2500
                self.log(f"Small–moderate dataset detected ({area:,.0f} km²). Consider using {suggested_spacing} m grid spacing.", Qgis.Info)
            elif 5_000 < area <= 20_000:
                suggested_spacing = 1000
                self.log(f"Small dataset detected ({area:,.0f} km²). Consider using {suggested_spacing} m grid spacing.", Qgis.Info)
            elif 0 < area <= 5_000:
                suggested_spacing = 500
                self.log(f"Very small dataset detected ({area:,.0f} km²). Consider using {suggested_spacing} m grid spacing.", Qgis.Info)
            else:
                suggested_spacing = None

            crs0 = hatar0.crs().authid()
            extent0 = hatar0.extent()
            
            grid_type = 2
            try:
                if hasattr(self.dlg, 'radioButton_diamond') and self.dlg.radioButton_diamond.isChecked():
                    grid_type = 3
                elif hasattr(self.dlg, 'radioButton_hexagon') and self.dlg.radioButton_hexagon.isChecked():
                    grid_type = 4
                else:
                    grid_type = 2
            except Exception:
                grid_type = 2

            xmin = extent0.xMinimum()
            xmax = extent0.xMaximum()
            ymin = extent0.yMinimum()
            ymax = extent0.yMaximum()

            if grid_type in (3, 4):
                pad_x = 0.5 * float(h_spacing)
                pad_y = 0.5 * float(v_spacing)

                xmin -= pad_x
                xmax += pad_x
                ymin -= pad_y
                ymax += pad_y

                w = xmax - xmin
                h = ymax - ymin

                def _snap_up(size, step):
                    step = float(step)
                    if step <= 0:
                        return size
                    rem = size % step
                    return size if rem == 0 else (size + (step - rem))

                w2 = _snap_up(w, h_spacing)
                h2 = _snap_up(h, v_spacing)

                xmax = xmin + w2
                ymin = ymax - h2

            extent_str = f"{xmin},{xmax},{ymin},{ymax}"
            create_res = processing.run("native:creategrid", {
                'TYPE': grid_type,
                'EXTENT': extent_str,
                'HSPACING': h_spacing,
                'VSPACING': v_spacing,
                'CRS': crs0,
                'OUTPUT': 'TEMPORARY_OUTPUT'
            })
            grid0_temp = create_res.get('OUTPUT')
            
            self._set_progress(progressbar, 2)
            
            self.log("Selecting grid cells that intersect boundary...", Qgis.Info)
            processing.run("native:extractbylocation", {
                'INPUT': grid0_temp,
                'PREDICATE': [0],
                'INTERSECT': boundary_path,
                'OUTPUT': grid0
            })
            
            addGrid0 = QgsVectorLayer(grid0, grid_name, "ogr")
            if not addGrid0.isValid():
                raise Exception("Grid creation failed!")
            
            grid_cell_count = addGrid0.featureCount()
            self.log(f"Grid created with {grid_cell_count} cells (whole cells that touch boundary)", Qgis.Info)
            
            self._set_progress(progressbar, 5)
            
            geol_layer = None
            if geology_path and geol_field:
                self.log("Processing geology data...", Qgis.Info)
                try:
                    geol_name = self._layer_base_name(geology_path, "geology")
                    layer11 = QgsVectorLayer(geology_path, geol_name, "ogr")
                    if layer11.isValid():
                        _add_layer(layer11, intermediate=True)
                        self._encode_unique_values(layer11, geol_field, out_field='r_value')
                        output_grid31 = self._output_gpkg_path(working_dir, geology_path, "_geology_grid", "geology")
                        geol_layer = self._vector_touch_variety(
                            grid_path=grid0,
                            poly_layer=layer11,
                            out_path=output_grid31,
                            out_field='_geol_variety',
                            value_field='r_value'
                        )
                        self._remove_field_if_exists(layer11, 'r_value')
                        if geol_layer and geol_layer.isValid():
                            geol_layer.setName(f"{geol_name}_geology_grid")
                        _add_layer(geol_layer, intermediate=True)
                except Exception as e:
                    self.log(f"Geology processing error (skipping): {str(e)}", Qgis.Warning)
            
            self._set_progress(progressbar, 20)
            
            pedo_layer = None
            if pedology_path and pedo_field:
                self.log("Processing pedology data...", Qgis.Info)
                try:
                    pedo_name = self._layer_base_name(pedology_path, "pedology")
                    layer12 = QgsVectorLayer(pedology_path, pedo_name, "ogr")
                    if layer12.isValid():
                        _add_layer(layer12, intermediate=True)
                        self._encode_unique_values(layer12, pedo_field, out_field='r_value')
                        output_grid32 = self._output_gpkg_path(working_dir, pedology_path, "_pedology_grid", "pedology")
                        pedo_layer = self._vector_touch_variety(
                            grid_path=grid0,
                            poly_layer=layer12,
                            out_path=output_grid32,
                            out_field='_pedo_variety',
                            value_field='r_value'
                        )
                        self._remove_field_if_exists(layer12, 'r_value')
                        if pedo_layer and pedo_layer.isValid():
                            pedo_layer.setName(f"{pedo_name}_pedology_grid")
                        _add_layer(pedo_layer, intermediate=True)
                except Exception as e:
                    self.log(f"Pedology processing error (skipping): {str(e)}", Qgis.Warning)
            
            self._set_progress(progressbar, 35)
            
            geom_layer = None
            stra_layer = None
            cut_dem2 = None
            compute_hydro = False
            if geom_source_idx == 0:
                compute_hydro = bool(dem_path)
            elif geom_source_idx == 1:
                try:
                    compute_hydro = bool(getattr(self.dlg, "checkBox_geomorphon_hydro", None) and self.dlg.checkBox_geomorphon_hydro.isChecked() and dem_path)
                except Exception:
                    compute_hydro = bool(dem_path)
            elif geom_source_idx == 2:
                compute_hydro = bool(dem_path)

            if geom_source_idx == 0:
                self.log("Processing vector geomorphology...", Qgis.Info)
                try:
                    geom_source_path = geom_line_path or geom_poly_path or geom_point_path
                    geom_base = self._layer_base_name(geom_source_path, "geomorphology")
                    output_grid33 = self._output_gpkg_path(working_dir, geom_source_path, "_geomorphology_grid", "geomorphology")
                    geom_inputs = [
                        (geom_line_path or None, geom_line_field or None, self._layer_base_name(geom_line_path, "geomorphology_lines")),
                        (geom_poly_path or None, geom_poly_field or None, self._layer_base_name(geom_poly_path, "geomorphology_polygons")),
                        (geom_point_path or None, geom_point_field or None, self._layer_base_name(geom_point_path, "geomorphology_points")),
                    ]
                    geom_layer = self._geomorphology_vector_variety(
                        grid_path=grid0,
                        out_path=output_grid33,
                        geom_inputs=geom_inputs
                    )
                    if geom_layer and geom_layer.isValid():
                        geom_layer.setName(f"{geom_base}_geomorphology_grid")
                        _add_layer(geom_layer, intermediate=True)
                except Exception as e:
                    self.log(f"Vector geomorphology failed: {str(e)}", Qgis.Warning)

            if geom_source_idx in (1, 2) or compute_hydro:
                if dem_path:
                    self.log("Preparing DEM for geomorphology/hydrography...", Qgis.Info)
                    try:
                        cut_dem2 = working_dir + "/cut_dem.tif"
                        crs2 = hatar0.crs().authid()
                        try:
                            processing.run("gdal:cliprasterbymasklayer", {
                                'INPUT': dem_path,
                                'MASK': boundary_path,
                                'SOURCE_CRS': crs2,
                                'TARGET_CRS': crs2,
                                'CROP_TO_CUTLINE': True,
                                'DATA_TYPE': 0,
                                'OUTPUT': cut_dem2
                            })
                        except Exception:
                            self.log("DEM clipping failed, trying alternative...", Qgis.Warning)
                            processing.run("gdal:warpreproject", {
                                'INPUT': dem_path,
                                'TARGET_CRS': crs2,
                                'OUTPUT': cut_dem2
                            })
                    except Exception as e:
                        self.log(f"DEM preparation failed: {str(e)}", Qgis.Warning)
                        cut_dem2 = None

            self._set_progress(progressbar, 45)

            if geom_source_idx == 1 and cut_dem2:
                self.log("Processing geomorphon from DEM...", Qgis.Info)
                geom2 = working_dir + "/geomorphon.tif"
                try:
                    search = int(self.dlg.spinBox_geomorphon_search.value()) if hasattr(self.dlg, 'spinBox_geomorphon_search') else 3
                    skip = int(self.dlg.spinBox_geomorphon_skip.value()) if hasattr(self.dlg, 'spinBox_geomorphon_skip') else 0
                    flat = float(self.dlg.doubleSpinBox_geomorphon_flat.value()) if hasattr(self.dlg, 'doubleSpinBox_geomorphon_flat') else 1.0
                    dist = int(self.dlg.spinBox_geomorphon_dist.value()) if hasattr(self.dlg, 'spinBox_geomorphon_dist') else 0
                    flag_m = bool(getattr(self.dlg, 'checkBox_geomorphon_m', None) and self.dlg.checkBox_geomorphon_m.isChecked())
                    flag_e = bool(getattr(self.dlg, 'checkBox_geomorphon_e', None) and self.dlg.checkBox_geomorphon_e.isChecked())
                    processing.run("grass7:r.geomorphon", {
                        'elevation': cut_dem2,
                        'search': search,
                        'skip': skip,
                        'flat': flat,
                        'dist': dist,
                        'forms': geom2,
                        '-m': flag_m,
                        '-e': flag_e
                    })

                    addRaster2 = QgsRasterLayer(geom2, "geomorphon_raster")
                    _add_layer(addRaster2, intermediate=True)

                    dem_base = self._layer_base_name(dem_path, "dem")
                    output_grid33 = self._output_gpkg_path(working_dir, dem_path, "_geomorphon_grid", "geomorphon")
                    processing.run("qgis:zonalstatisticsfb", {
                        'INPUT': grid0,
                        'INPUT_RASTER': geom2,
                        'RASTER_BAND': 1,
                        'COLUMN_PREFIX': '_geom_',
                        'STATISTICS': [1, 2, 3, 4, 5, 6, 7, 8, 9, 10],
                        'OUTPUT': output_grid33
                    })

                    geom_layer = QgsVectorLayer(output_grid33, f"{dem_base}_geomorphon_grid", "ogr")
                    if geom_layer.isValid() and geom_layer.fields().count() >= 15:
                        geom_layer.dataProvider().deleteAttributes([6, 7, 8, 9, 10, 11, 12, 13, 14])
                        geom_layer.updateFields()
                    _add_layer(geom_layer, intermediate=True)
                except Exception as e:
                    self.log(f"Geomorphon failed: {str(e)}", Qgis.Warning)
            elif geom_source_idx == 2 and geomorphon_raster_path:
                self.log("Processing uploaded geomorphon raster...", Qgis.Info)
                try:
                    geom_name = self._layer_base_name(geomorphon_raster_path, "geomorphon")
                    geom_raster = QgsRasterLayer(geomorphon_raster_path, geom_name)
                    if not geom_raster.isValid():
                        raise Exception("Uploaded geomorphon raster is invalid.")
                    _add_layer(geom_raster, intermediate=True)
                    output_grid33 = self._output_gpkg_path(working_dir, geomorphon_raster_path, "_geomorphon_grid", "geomorphon")
                    processing.run("qgis:zonalstatisticsfb", {
                        'INPUT': grid0,
                        'INPUT_RASTER': geomorphon_raster_path,
                        'RASTER_BAND': 1,
                        'COLUMN_PREFIX': '_geom_',
                        'STATISTICS': [1, 2, 3, 4, 5, 6, 7, 8, 9, 10],
                        'OUTPUT': output_grid33
                    })
                    geom_layer = QgsVectorLayer(output_grid33, f"{geom_name}_geomorphon_grid", "ogr")
                    if geom_layer.isValid() and geom_layer.fields().count() >= 15:
                        geom_layer.dataProvider().deleteAttributes([6, 7, 8, 9, 10, 11, 12, 13, 14])
                        geom_layer.updateFields()
                    _add_layer(geom_layer, intermediate=True)
                except Exception as e:
                    self.log(f"Uploaded geomorphon raster processing failed: {str(e)}", Qgis.Warning)

            self._set_progress(progressbar, 55)

            if compute_hydro and cut_dem2:
                self.log("Processing hydrography from DEM (Strahler)...", Qgis.Info)
                try:
                    filled_output5 = working_dir + "/filled_dem.sdat"
                    flow_dir_output = 'TEMPORARY_OUTPUT'
                    watershed_output = 'TEMPORARY_OUTPUT'

                    try:
                        self.log("Filling DEM sinks and calculating flow direction...", Qgis.Info)
                        processing.run("sagang:fillsinkswangliu", {
                            'ELEV': cut_dem2,
                            'FILLED': filled_output5,
                            'FDIR': flow_dir_output,
                            'WSHED': watershed_output,
                            'MINSLOPE': 0.1
                        })
                    except Exception:
                        self.log("Wang & Liu fill failed, using simple fill...", Qgis.Warning)
                        processing.run("saga:fillsinks", {
                            'DEM': cut_dem2,
                            'RESULT': filled_output5
                        })

                    self._set_progress(progressbar, 60)

                    strahler_output5 = working_dir + "/strahler.sdat"
                    processing.run("sagang:strahlerorder", {
                        'DEM': filled_output5,
                        'STRAHLER': strahler_output5
                    })

                    self._set_progress(progressbar, 65)

                    dem_name = self._layer_base_name(dem_path, "dem")
                    output_grid5 = self._output_gpkg_path(working_dir, dem_path, "_strahler_grid", "strahler")
                    processing.run("qgis:zonalstatisticsfb", {
                        'INPUT': grid0,
                        'INPUT_RASTER': strahler_output5,
                        'RASTER_BAND': 1,
                        'COLUMN_PREFIX': '_stra_',
                        'STATISTICS': [6],
                        'OUTPUT': output_grid5
                    })

                    stra_layer = QgsVectorLayer(output_grid5, f"{dem_name}_strahler_grid", "ogr")
                    stra_layer.updateFields()
                    _add_layer(stra_layer, intermediate=True)

                    with edit(stra_layer):
                        for feature in stra_layer.getFeatures():
                            val = feature['_stra_max']
                            if val is None:
                                val = 0
                            new_value = math.ceil(val / 2)
                            feature.setAttribute(feature.fieldNameIndex('_stra_max'), new_value)
                            stra_layer.updateFeature(feature)
                except Exception as e:
                    self.log(f"Hydrography/Strahler processing error (skipping): {str(e)}", Qgis.Warning)
            self._set_progress(progressbar, 70)
            
            mine_layer = None
            if mineral_path and mineral_field:
                self.log("Processing mineralogy data...", Qgis.Info)
                try:
                    mineral_name = self._layer_base_name(mineral_path, "mineral")
                    layer61 = QgsVectorLayer(mineral_path, mineral_name, "ogr")
                    if layer61.isValid():
                        _add_layer(layer61, intermediate=True)
                        self._encode_unique_values(layer61, mineral_field, out_field='r_value')
                        mineral_grid = self._output_gpkg_path(working_dir, mineral_path, "_mineral_grid", "mineral")
                        processing.run("native:countpointsinpolygon", {
                            'POLYGONS': grid0,
                            'POINTS': layer61,
                            'CLASSFIELD': 'r_value',
                            'FIELD': '_mineral_idx',
                            'OUTPUT': mineral_grid
                        })
                        mine_layer = QgsVectorLayer(mineral_grid, f'{mineral_name}_mineral_grid', "ogr")
                        _add_layer(mine_layer, intermediate=True)
                except Exception as e:
                    self.log(f"Mineralogy error (skipping): {str(e)}", Qgis.Warning)
            
            self._set_progress(progressbar, 80)
            
            foss_layer = None
            if palaeo_path and palaeo_field:
                self.log("Processing palaeontology data...", Qgis.Info)
                try:
                    palaeo_name = self._layer_base_name(palaeo_path, "palaeo")
                    layer62 = QgsVectorLayer(palaeo_path, palaeo_name, "ogr")
                    if layer62.isValid():
                        _add_layer(layer62, intermediate=True)
                        self._encode_unique_values(layer62, palaeo_field, out_field='r_value')
                        fossil_grid = self._output_gpkg_path(working_dir, palaeo_path, "_palaeontology_grid", "palaeo")
                        processing.run("native:countpointsinpolygon", {
                            'POLYGONS': grid0,
                            'POINTS': layer62,
                            'CLASSFIELD': 'r_value',
                            'FIELD': '_fossil_idx',
                            'OUTPUT': fossil_grid
                        })
                        foss_layer = QgsVectorLayer(fossil_grid, f'{palaeo_name}_palaeontology_grid', "ogr")
                        _add_layer(foss_layer, intermediate=True)
                except Exception as e:
                    self.log(f"Palaeontology error (skipping): {str(e)}", Qgis.Warning)
            
            self._set_progress(progressbar, 85)
            
            if lakes_path:
                self.log("Processing lake/sea data...", Qgis.Info)
                try:
                    lakes_name = self._layer_base_name(lakes_path, "lakes")
                    lakes7_input = QgsVectorLayer(lakes_path, lakes_name, "ogr")
                    if lakes7_input.isValid():
                        newField7 = QgsField('_lakes', QVariant.Int)
                        addGrid0.dataProvider().addAttributes([newField7])
                        addGrid0.updateFields()
                        
                        selection = processing.run("native:selectbylocation", {
                            'INPUT': addGrid0,
                            'INTERSECT': lakes7_input,
                            'METHOD': 0,
                            'PREDICATE': [0]
                        })
                        
                        with edit(addGrid0):
                            for id in selection['OUTPUT'].selectedFeatureIds():
                                feature = addGrid0.getFeature(id)
                                feature['_lakes'] = 3
                                addGrid0.updateFeature(feature)
                        
                        addGrid0.removeSelection()
                except Exception as e:
                    self.log(f"Lakes error (skipping): {str(e)}", Qgis.Warning)
            
            self._set_progress(progressbar, 87)
            
            self.log("Joining thematic fields to grid...", Qgis.Info)
            grid = addGrid0
            
            if geol_layer:
                try:
                    joinObject1 = QgsVectorLayerJoinInfo()
                    joinObject1.setJoinFieldName('id')
                    joinObject1.setTargetFieldName('id')
                    joinObject1.setJoinLayerId(geol_layer.id())
                    joinObject1.setUsingMemoryCache(True)
                    joinObject1.setJoinLayer(geol_layer)
                    joinObject1.setPrefix('J')
                    joinObject1.setJoinFieldNamesSubset(['_geol_variety'])
                    grid.addJoin(joinObject1)
                except Exception as e:
                    self.log(f"Geology join failed: {str(e)}", Qgis.Warning)
            
            if pedo_layer:
                try:
                    joinObject2 = QgsVectorLayerJoinInfo()
                    joinObject2.setJoinFieldName('id')
                    joinObject2.setTargetFieldName('id')
                    joinObject2.setJoinLayerId(pedo_layer.id())
                    joinObject2.setUsingMemoryCache(True)
                    joinObject2.setJoinLayer(pedo_layer)
                    joinObject2.setPrefix('J')
                    joinObject2.setJoinFieldNamesSubset(['_pedo_variety'])
                    grid.addJoin(joinObject2)
                except Exception as e:
                    self.log(f"Pedology join failed: {str(e)}", Qgis.Warning)
            
            if geom_layer:
                try:
                    joinObject3 = QgsVectorLayerJoinInfo()
                    joinObject3.setJoinFieldName('id')
                    joinObject3.setTargetFieldName('id')
                    joinObject3.setJoinLayerId(geom_layer.id())
                    joinObject3.setUsingMemoryCache(True)
                    joinObject3.setJoinLayer(geom_layer)
                    joinObject3.setPrefix('J')
                    joinObject3.setJoinFieldNamesSubset(['_geom_variety'])
                    grid.addJoin(joinObject3)
                except Exception as e:
                    self.log(f"Geomorphon join failed: {str(e)}", Qgis.Warning)
            
            if stra_layer:
                try:
                    joinObject4 = QgsVectorLayerJoinInfo()
                    joinObject4.setJoinFieldName('id')
                    joinObject4.setTargetFieldName('id')
                    joinObject4.setJoinLayerId(stra_layer.id())
                    joinObject4.setUsingMemoryCache(True)
                    joinObject4.setJoinLayer(stra_layer)
                    joinObject4.setPrefix('J')
                    joinObject4.setJoinFieldNamesSubset(['_stra_max'])
                    grid.addJoin(joinObject4)
                except Exception as e:
                    self.log(f"Strahler join failed: {str(e)}", Qgis.Warning)
            
            if mine_layer:
                try:
                    joinObject5 = QgsVectorLayerJoinInfo()
                    joinObject5.setJoinFieldName('id')
                    joinObject5.setTargetFieldName('id')
                    joinObject5.setJoinLayerId(mine_layer.id())
                    joinObject5.setUsingMemoryCache(True)
                    joinObject5.setJoinLayer(mine_layer)
                    joinObject5.setPrefix('J')
                    joinObject5.setJoinFieldNamesSubset(['_mineral_idx'])
                    grid.addJoin(joinObject5)
                except Exception as e:
                    self.log(f"Mineralogy join failed: {str(e)}", Qgis.Warning)
            
            if foss_layer:
                try:
                    joinObject6 = QgsVectorLayerJoinInfo()
                    joinObject6.setJoinFieldName('id')
                    joinObject6.setTargetFieldName('id')
                    joinObject6.setJoinLayerId(foss_layer.id())
                    joinObject6.setUsingMemoryCache(True)
                    joinObject6.setJoinLayer(foss_layer)
                    joinObject6.setPrefix('J')
                    joinObject6.setJoinFieldNamesSubset(['_fossil_idx'])
                    grid.addJoin(joinObject6)
                except Exception as e:
                    self.log(f"Palaeontology join failed: {str(e)}", Qgis.Warning)
            
            self._set_progress(progressbar, 90)
            
            if geol_layer:
                try:
                    with edit(grid):
                        features_to_delete = []
                        for feature in grid.getFeatures():
                            geol_val = feature['J_geol_variety']
                            if geol_val is None:
                                features_to_delete.append(feature.id())
                        
                        if features_to_delete:
                            grid.deleteFeatures(features_to_delete)
                            self.log(f"Deleted {len(features_to_delete)} cells with NULL geology", Qgis.Info)
                except Exception as e:
                    self.log(f"NULL deletion warning: {str(e)}", Qgis.Warning)
            
            self._set_progress(progressbar, 93)

            try:
                do_norm = bool(getattr(self.dlg, "checkBox_normalize", None) and self.dlg.checkBox_normalize.isChecked())
            except Exception:
                do_norm = False

            if do_norm:
                self.log("Normalizing subindices (0-1)...", Qgis.Info)
                self._add_normalized_fields(grid)

            self.log("Calculating final geodiversity index...", Qgis.Info)
            prov = grid.dataProvider()
            newField8 = QgsField('_GEODIV', QVariant.Int)
            prov.addAttributes([newField8])
            grid.updateFields()
            
            idx = grid.fields().lookupField('_GEODIV')
            context = QgsExpressionContext()
            expression = QgsExpression(
                '(CASE WHEN "_lakes" IS NOT NULL THEN "_lakes" ELSE 0 END) + '
                '(CASE WHEN "J_geol_variety" IS NOT NULL THEN "J_geol_variety" ELSE 0 END) + '
                '(CASE WHEN "J_pedo_variety" IS NOT NULL THEN "J_pedo_variety" ELSE 0 END) + '
                '(CASE WHEN "J_geom_variety" IS NOT NULL THEN "J_geom_variety" ELSE 0 END) + '
                '(CASE WHEN "J_stra_max" IS NOT NULL THEN "J_stra_max" ELSE 0 END) + '
                '(CASE WHEN "J_mineral_idx" IS NOT NULL THEN "J_mineral_idx" ELSE 0 END) + '
                '(CASE WHEN "J_fossil_idx" IS NOT NULL THEN "J_fossil_idx" ELSE 0 END)'
            )
            
            scope = QgsExpressionContextScope()
            scope.setFields(grid.fields())
            context.appendScope(scope)
            expression.prepare(context)
            
            with edit(grid):
                for feature in grid.getFeatures():
                    context.setFeature(feature)
                    geodiv = expression.evaluate(context)
                    atts = {idx: geodiv}
                    grid.dataProvider().changeAttributeValues({feature.id(): atts})
            
            self._set_progress(progressbar, 97)
            
            _add_layer(addGrid0, intermediate=False)

            try:
                style_field = "N_sum" if ("do_norm" in locals() and do_norm) else "_GEODIV"
                self._apply_output_style(addGrid0, style_field)
            except Exception:
                pass
            
            output_files = []
            manifest_path = working_dir + "/" + grid_name + "_MANIFEST.txt"
            
            for file in os.listdir(working_dir):
                if os.path.isfile(os.path.join(working_dir, file)):
                    file_size_mb = os.path.getsize(os.path.join(working_dir, file)) / (1024 * 1024)
                    output_files.append(f"{file} ({file_size_mb:.2f} MB)")
            
            with open(manifest_path, 'w', encoding='utf-8') as f:
                f.write("=" * 70 + "\n")
                f.write("GEODIVERSITY CALCULATOR v2.1 - OUTPUT FILE MANIFEST\n")
                f.write("=" * 70 + "\n\n")
                f.write(f"Analysis Date: {QDateTime.currentDateTime().toString('yyyy-MM-dd HH:mm:ss')}\n")
                f.write(f"Boundary Area: {boundary_area_km2:.0f} km²\n")
                f.write(f"Grid Spacing: {h_spacing}m x {v_spacing}m\n")
                f.write(f"Grid Cells: {grid_cell_count}\n")
                f.write(f"Working Directory: {working_dir}\n\n")
                f.write("=" * 70 + "\n")
                f.write("OUTPUT FILES:\n")
                f.write("=" * 70 + "\n\n")
                for file in sorted(output_files):
                    f.write(f"  • {file}\n")
                f.write("\n" + "=" * 70 + "\n")
                f.write("ANALYSIS COMPONENTS:\n")
                f.write("=" * 70 + "\n\n")
                if geology_path:
                    f.write("  ✓ Geology (variety)\n")
                if pedology_path:
                    f.write("  ✓ Pedology (variety)\n")
                if dem_path:
                    f.write("  ✓ DEM (geomorphon + Strahler stream order)\n")
                    f.write("  ✓ Flow direction & watershed\n")
                if lakes_path:
                    f.write("  ✓ Lakes/water bodies\n")
                if mineral_path:
                    f.write("  ✓ Mineralogy (occurrence diversity)\n")
                if palaeo_path:
                    f.write("  ✓ Palaeontology (fossil diversity)\n")
                f.write("\n" + "=" * 70 + "\n")
                f.write(f"Final Grid: {grid_name}.gpkg\n")
                f.write(f"Total Files: {len(output_files)}\n")
                f.write("=" * 70 + "\n")
            
            self.log(f"Output manifest saved: {manifest_path}", Qgis.Info)
            self._set_progress(progressbar, 100)
            
            self.iface.messageBar().pushMessage(
                "Success",
                f"GeoDiversity v2.1 completed! {grid_cell_count} cells, {len(output_files)} files created. See {grid_name}_MANIFEST.txt",
                level=Qgis.Success,
                duration=20
            )
            
            self.log("Analysis complete!", Qgis.Info)
        
        except Exception as e:
            error_msg = f"GeoDiversity v2.1 Error: {str(e)}"
            self.log(error_msg, Qgis.Critical)
            self.log(traceback.format_exc(), Qgis.Critical)
            self.iface.messageBar().pushMessage("Error", error_msg, level=Qgis.Critical, duration=15)
        
        finally:
            if progressbar:
                self.iface.messageBar().clearWidgets()
