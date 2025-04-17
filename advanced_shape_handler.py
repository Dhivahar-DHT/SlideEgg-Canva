import xml.etree.ElementTree as ET
from zipfile import ZipFile
import io
import base64
import math

class AdvancedShapeHandler:
    def __init__(self):
        self.namespace = {
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'
        }
    
    def extract_shape_properties(self, shape):
        """Extract advanced shape properties directly from OOXML"""
        try:
            if not hasattr(shape, 'element'):
                return None
                
            props = {
                'effects': [],
                'gradient': None,
                'custom_geometry': None,
                'shadow': None,
                'reflection': None,
                'glow': None,
                'soft_edges': None
            }
            
            # Get shape properties element
            sp_pr = shape.element.find('.//a:spPr', self.namespace)
            if sp_pr is None:
                return None
                
            # Extract gradient fill
            grad_fill = sp_pr.find('.//a:gradFill', self.namespace)
            if grad_fill is not None:
                props['gradient'] = self._extract_gradient(grad_fill)
            
            # Extract custom geometry
            custom_geom = sp_pr.find('.//a:custGeom', self.namespace)
            if custom_geom is not None:
                props['custom_geometry'] = self._extract_custom_geometry(custom_geom)
            
            # Extract effects
            effects = sp_pr.find('.//a:effectLst', self.namespace)
            if effects is not None:
                props['effects'] = self._extract_effects(effects)
            
            return props
            
        except Exception as e:
            print(f"Error extracting advanced properties: {e}")
            return None
    
    def _extract_gradient(self, grad_fill):
        """Extract gradient information"""
        try:
            gradient = {
                'type': 'linear',  # default
                'angle': 0,
                'stops': []
            }
            
            # Get gradient type
            if grad_fill.find('.//a:lin', self.namespace) is not None:
                lin = grad_fill.find('.//a:lin', self.namespace)
                angle = int(lin.get('ang', '0')) / 60000  # Convert to degrees
                gradient['angle'] = angle
            elif grad_fill.find('.//a:path', self.namespace) is not None:
                gradient['type'] = 'radial'
            
            # Get gradient stops
            gs_list = grad_fill.findall('.//a:gs', self.namespace)
            for gs in gs_list:
                pos = int(gs.get('pos', '0')) / 100000  # Normalize to 0-1
                
                # Get color
                color = None
                srgb_clr = gs.find('.//a:srgbClr', self.namespace)
                if srgb_clr is not None:
                    color = f"#{srgb_clr.get('val')}"
                
                if color:
                    gradient['stops'].append({
                        'offset': pos,
                        'color': color
                    })
            
            return gradient
            
        except Exception as e:
            print(f"Error extracting gradient: {e}")
            return None
    
    def _extract_custom_geometry(self, custom_geom):
        """Extract custom shape geometry"""
        try:
            geometry = {
                'paths': [],
                'rect': None
            }
            
            # Get shape boundaries
            rect = custom_geom.find('.//a:rect', self.namespace)
            if rect is not None:
                geometry['rect'] = {
                    'l': int(rect.get('l', '0')) / 100000,
                    't': int(rect.get('t', '0')) / 100000,
                    'r': int(rect.get('r', '0')) / 100000,
                    'b': int(rect.get('b', '0')) / 100000
                }
            
            # Get path list
            path_list = custom_geom.find('.//a:pathLst', self.namespace)
            if path_list is not None:
                for path in path_list.findall('.//a:path', self.namespace):
                    path_data = []
                    
                    # Process each command
                    for cmd in path:
                        tag = cmd.tag.split('}')[-1]
                        if tag == 'moveTo':
                            pt = cmd.find('.//a:pt', self.namespace)
                            x = int(pt.get('x', '0')) / 100000
                            y = int(pt.get('y', '0')) / 100000
                            path_data.append(f'M {x} {y}')
                        elif tag == 'lnTo':
                            pt = cmd.find('.//a:pt', self.namespace)
                            x = int(pt.get('x', '0')) / 100000
                            y = int(pt.get('y', '0')) / 100000
                            path_data.append(f'L {x} {y}')
                        elif tag == 'cubicBezTo':
                            pts = cmd.findall('.//a:pt', self.namespace)
                            if len(pts) == 3:
                                x1 = int(pts[0].get('x', '0')) / 100000
                                y1 = int(pts[0].get('y', '0')) / 100000
                                x2 = int(pts[1].get('x', '0')) / 100000
                                y2 = int(pts[1].get('y', '0')) / 100000
                                x3 = int(pts[2].get('x', '0')) / 100000
                                y3 = int(pts[2].get('y', '0')) / 100000
                                path_data.append(f'C {x1} {y1} {x2} {y2} {x3} {y3}')
                        elif tag == 'arcTo':
                            # Convert arc to cubic bezier curves
                            # This is a simplified version - you might need more complex arc handling
                            wR = int(cmd.get('wR', '0')) / 100000
                            hR = int(cmd.get('hR', '0')) / 100000
                            stAng = int(cmd.get('stAng', '0')) / 60000
                            swAng = int(cmd.get('swAng', '0')) / 60000
                            path_data.append(f'A {wR} {hR} {stAng} {swAng > 180} {swAng > 0}')
                        elif tag == 'close':
                            path_data.append('Z')
                    
                    if path_data:
                        geometry['paths'].append(' '.join(path_data))
            
            return geometry
            
        except Exception as e:
            print(f"Error extracting custom geometry: {e}")
            return None
    
    def _extract_effects(self, effects):
        """Extract shape effects"""
        try:
            effect_list = []
            
            # Extract shadow
            shadow = effects.find('.//a:outerShdw', self.namespace)
            if shadow is not None:
                effect_list.append({
                    'type': 'shadow',
                    'color': self._get_effect_color(shadow),
                    'opacity': int(shadow.get('alpha', '100000')) / 100000,
                    'blur': int(shadow.get('blurRad', '0')) / 12700,
                    'offset': {
                        'x': int(shadow.get('dx', '0')) / 12700,
                        'y': int(shadow.get('dy', '0')) / 12700
                    }
                })
            
            # Extract glow
            glow = effects.find('.//a:glow', self.namespace)
            if glow is not None:
                effect_list.append({
                    'type': 'glow',
                    'color': self._get_effect_color(glow),
                    'opacity': int(glow.get('alpha', '100000')) / 100000,
                    'radius': int(glow.get('rad', '0')) / 12700
                })
            
            # Extract soft edges
            soft = effects.find('.//a:softEdge', self.namespace)
            if soft is not None:
                effect_list.append({
                    'type': 'soft-edge',
                    'radius': int(soft.get('rad', '0')) / 12700
                })
            
            return effect_list
            
        except Exception as e:
            print(f"Error extracting effects: {e}")
            return []
    
    def _get_effect_color(self, effect_elem):
        """Extract color from effect element"""
        try:
            srgb_clr = effect_elem.find('.//a:srgbClr', self.namespace)
            if srgb_clr is not None:
                return f"#{srgb_clr.get('val')}"
            return '#000000'  # Default black
        except Exception as e:
            print(f"Error getting effect color: {e}")
            return '#000000'
    
    def convert_to_fabric(self, shape_props):
        """Convert PowerPoint shape properties to Fabric.js format"""
        try:
            if not shape_props:
                return {}
                
            fabric_props = {}
            
            # Convert gradient
            if shape_props.get('gradient'):
                grad = shape_props['gradient']
                fabric_props['fill'] = {
                    'type': grad['type'],
                    'coords': {
                        'x1': 0,
                        'y1': 0,
                        'x2': 0,
                        'y2': 1
                    } if grad['type'] == 'linear' else {
                        'x1': 0.5,
                        'y1': 0.5,
                        'r1': 0,
                        'x2': 0.5,
                        'y2': 0.5,
                        'r2': 0.5
                    },
                    'colorStops': {
                        str(stop['offset']): stop['color']
                        for stop in grad['stops']
                    }
                }
                
                if grad['type'] == 'linear':
                    # Convert angle to coordinates
                    angle = grad['angle'] * Math.PI / 180
                    fabric_props['fill']['coords'] = {
                        'x1': 0.5 - 0.5 * Math.cos(angle),
                        'y1': 0.5 - 0.5 * Math.sin(angle),
                        'x2': 0.5 + 0.5 * Math.cos(angle),
                        'y2': 0.5 + 0.5 * Math.sin(angle)
                    }
            
            # Convert custom geometry
            if shape_props.get('custom_geometry'):
                geom = shape_props['custom_geometry']
                if geom['paths']:
                    fabric_props['path'] = geom['paths'][0]  # Use first path
                if geom['rect']:
                    fabric_props['clipPath'] = {
                        'type': 'rect',
                        'left': geom['rect']['l'],
                        'top': geom['rect']['t'],
                        'width': geom['rect']['r'] - geom['rect']['l'],
                        'height': geom['rect']['b'] - geom['rect']['t']
                    }
            
            # Convert effects
            for effect in shape_props.get('effects', []):
                if effect['type'] == 'shadow':
                    fabric_props['shadow'] = {
                        'color': effect['color'],
                        'blur': effect['blur'],
                        'offsetX': effect['offset']['x'],
                        'offsetY': effect['offset']['y'],
                        'opacity': effect['opacity']
                    }
                elif effect['type'] == 'glow':
                    # Fabric.js doesn't support glow directly
                    # We can simulate it with multiple shadows
                    fabric_props['shadow'] = {
                        'color': effect['color'],
                        'blur': effect['radius'] * 2,
                        'offsetX': 0,
                        'offsetY': 0,
                        'opacity': effect['opacity']
                    }
                elif effect['type'] == 'soft-edge':
                    # Simulate soft edges with a very slight blur
                    fabric_props['shadow'] = {
                        'color': '#000000',
                        'blur': effect['radius'],
                        'offsetX': 0,
                        'offsetY': 0,
                        'opacity': 0.1
                    }
            
            return fabric_props
            
        except Exception as e:
            print(f"Error converting to Fabric.js format: {e}")
            return {} 