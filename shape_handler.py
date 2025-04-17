from pptx.enum.shapes import MSO_SHAPE_TYPE, MSO_SHAPE
from pptx.enum.dml import MSO_FILL
import math
from advanced_shape_handler import AdvancedShapeHandler

class ShapeHandler:
    def __init__(self, color_handler, text_handler):
        self.color_handler = color_handler
        self.text_handler = text_handler
        self.advanced_handler = AdvancedShapeHandler()
    
    def process_shape(self, shape):
        """Process any type of shape and return its properties"""
        try:
            print(f"\nProcessing shape: {shape.shape_type}")  # Debug log
            base_props = self._get_base_properties(shape)
            
            # Get advanced properties first
            advanced_props = self.advanced_handler.extract_shape_properties(shape)
            if advanced_props:
                fabric_props = self.advanced_handler.convert_to_fabric(advanced_props)
                base_props.update(fabric_props)
            
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                return self._process_group(shape, base_props)
            elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                return self._process_picture(shape, base_props)
            elif shape.shape_type == MSO_SHAPE_TYPE.TEXT_BOX:
                return self._process_textbox(shape, base_props)
            elif shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
                return self._process_autoshape(shape, base_props)
            elif shape.shape_type == MSO_SHAPE_TYPE.FREEFORM:
                return self._process_freeform(shape, base_props)
            elif shape.shape_type == MSO_SHAPE_TYPE.LINE:
                return self._process_line(shape, base_props)
            else:
                print(f"Unsupported shape type: {shape.shape_type}")
                return None
                
        except Exception as e:
            print(f"Error processing shape: {e}")
            return None
    
    def _get_base_properties(self, shape):
        """Get basic properties common to all shapes"""
        return {
            "left": shape.left / 12700,  # Convert EMU to points
            "top": shape.top / 12700,
            "width": shape.width / 12700,
            "height": shape.height / 12700,
            "rotation": shape.rotation if hasattr(shape, 'rotation') else 0
        }
    
    def _process_group(self, group_shape, base_props):
        """Process a group of shapes"""
        shapes_data = []
        for shape in group_shape.shapes:
            shape_data = self.process_shape(shape)
            if shape_data:
                shapes_data.append(shape_data)
        
        if shapes_data:
            return {
                **base_props,
                "type": "group",
                "objects": shapes_data
            }
        return None
    
    def _process_picture(self, shape, base_props):
        """Process a picture shape"""
        try:
            print(f"Processing picture shape: {shape}")  # Debug log
            if hasattr(shape, 'image'):
                import base64
                image_data = base64.b64encode(shape.image.blob).decode()
                image_type = shape.image.content_type.split('/')[-1]
                
                # Don't try to access fill property for pictures
                opacity = 1
                if hasattr(shape, 'element'):
                    # Try to get opacity from element properties if available
                    alpha_mod = shape.element.find('.//a:alphaModFix', 
                        {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'})
                    if alpha_mod is not None:
                        amt = alpha_mod.get('amt')
                        if amt:
                            opacity = int(amt) / 100000
                
                return {
                    **base_props,
                    "type": "image",
                    "src": f"data:image/{image_type};base64,{image_data}",
                    "opacity": opacity,
                    "crossOrigin": "anonymous"  # Add this to handle CORS issues
                }
        except Exception as e:
            print(f"Error processing picture: {e}")
            print(f"Picture properties: {dir(shape)}")  # Debug log
        return None
    
    def _process_textbox(self, shape, base_props):
        """Process a textbox shape"""
        try:
            # Get text properties
            text_props = self.text_handler.get_text_properties(shape)
            fabric_text = self.text_handler.convert_to_fabric_text(text_props)
            
            if fabric_text:
                return {
                    **base_props,
                    **fabric_text,
                    "backgroundColor": self.color_handler.get_shape_color(shape)
                }
        except Exception as e:
            print(f"Error processing textbox: {e}")
        return None
    
    def _process_autoshape(self, shape, base_props):
        """Process an auto shape"""
        try:
            shape_type = "rect"  # Default to rectangle
            
            # Handle special shapes
            if hasattr(shape, 'auto_shape_type'):
                # Handle triangles (often used for markers)
                if shape.auto_shape_type == MSO_SHAPE.ISOSCELES_TRIANGLE:
                    shape_type = "triangle"
                # Add more shape type mappings as needed
            
            shape_data = {
                **base_props,
                "type": shape_type
            }
            
            # Get fill color
            fill_color = self.color_handler.get_shape_color(shape)
            if fill_color:
                shape_data["fill"] = fill_color
            else:
                print("No fill color found for shape")  # Debug log
            
            # Get line properties
            line_props = self._get_line_properties(shape)
            if not line_props.get('stroke'):
                line_props['stroke'] = '#000000'  # Default black if no color specified
            if not line_props.get('strokeWidth'):
                line_props['strokeWidth'] = 1  # Default width if not specified
            
            shape_data.update(line_props)
            
            # If it's a triangle marker, adjust the rotation
            if shape_type == "triangle":
                shape_data["angle"] = (shape_data.get("rotation", 0) + 180) % 360  # Point downward
            
            return shape_data
            
        except Exception as e:
            print(f"Error processing auto shape: {e}")
            return None
    
    def _process_freeform(self, shape, base_props):
        """Process a freeform shape"""
        try:
            path_data = self._get_path_data(shape)
            if path_data:
                shape_data = {
                    **base_props,
                    "type": "path",
                    "path": path_data
                }
                
                # Get fill color
                fill_color = self.color_handler.get_shape_color(shape)
                if fill_color:
                    shape_data["fill"] = fill_color
                
                # Get line properties
                line_props = self._get_line_properties(shape)
                shape_data.update(line_props)
                
                return shape_data
                
        except Exception as e:
            print(f"Error processing freeform shape: {e}")
        return None
    
    def _process_line(self, shape, base_props):
        """Process a line shape"""
        try:
            # Calculate line endpoints
            start_x = shape.left / 12700
            start_y = shape.top / 12700
            end_x = (shape.left + shape.width) / 12700
            end_y = (shape.top + shape.height) / 12700
            
            # Create path data for the line
            path_data = f'M {start_x} {start_y} L {end_x} {end_y}'
            
            shape_data = {
                **base_props,
                "type": "path",
                "path": path_data,
                "fill": "transparent"  # Lines don't have fill
            }
            
            # Get line properties
            line_props = self._get_line_properties(shape)
            if not line_props.get('stroke'):
                line_props['stroke'] = '#000000'  # Default black if no color specified
            if not line_props.get('strokeWidth'):
                line_props['strokeWidth'] = 1  # Default width if not specified
            
            shape_data.update(line_props)
            
            return shape_data
            
        except Exception as e:
            print(f"Error processing line shape: {e}")
            return None
    
    def _get_line_properties(self, shape):
        """Get line properties of a shape"""
        props = {}
        try:
            if hasattr(shape, 'line'):
                line = shape.line
                if hasattr(line, 'color') and line.color:
                    color = self.color_handler.get_shape_color(line)
                    if color:
                        props['stroke'] = color
                
                if hasattr(line, 'width'):
                    props['strokeWidth'] = line.width / 12700  # Convert EMU to points
        except Exception as e:
            print(f"Error getting line properties: {e}")
        return props
    
    def _get_path_data(self, shape):
        """Extract path data from a shape"""
        try:
            if hasattr(shape, 'element'):
                path_elem = shape.element.find('.//a:path', 
                    {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'})
                if path_elem is not None:
                    return self._extract_path_commands(path_elem)
        except Exception as e:
            print(f"Error getting path data: {e}")
        return None
    
    def _extract_path_commands(self, path_elem):
        """Extract path commands from XML element"""
        try:
            commands = []
            current_x = 0
            current_y = 0
            
            for child in path_elem:
                tag = child.tag.split('}')[-1]
                
                if tag == 'moveTo':
                    pt = child.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}pt')
                    x = float(pt.get('x')) / 12700
                    y = float(pt.get('y')) / 12700
                    commands.append(f'M {x} {y}')
                    current_x, current_y = x, y
                    
                elif tag == 'lnTo':
                    pt = child.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}pt')
                    x = float(pt.get('x')) / 12700
                    y = float(pt.get('y')) / 12700
                    commands.append(f'L {x} {y}')
                    current_x, current_y = x, y
                    
                elif tag == 'cubicBezTo':
                    pts = child.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/main}pt')
                    if len(pts) == 3:
                        x1 = float(pts[0].get('x')) / 12700
                        y1 = float(pts[0].get('y')) / 12700
                        x2 = float(pts[1].get('x')) / 12700
                        y2 = float(pts[1].get('y')) / 12700
                        x3 = float(pts[2].get('x')) / 12700
                        y3 = float(pts[2].get('y')) / 12700
                        commands.append(f'C {x1} {y1} {x2} {y2} {x3} {y3}')
                        current_x, current_y = x3, y3
                    
                elif tag == 'close':
                    commands.append('Z')
            
            return ' '.join(commands)
            
        except Exception as e:
            print(f"Error extracting path commands: {e}")
            return None 