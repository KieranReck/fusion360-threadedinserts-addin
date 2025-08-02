#Author-
#Description- ThreadedInsertAddin - Combined add-in for threaded insert parameter management and appearance application

import adsk.core
import adsk.fusion
import traceback

# Global variables to maintain references
app = None
ui = None
handlers = []

# --------- CONFIGURATION ----------------------------------------------------
# Configuration: Choose which tab to place the Threaded Inserts panel
# Uncomment ONE of the following lines, do not leave any leading spaces/indents:

#TARGET_TAB_ID = 'SolidTab'      # Places panel in SOLID tab (main design commands)
TARGET_TAB_ID = 'ToolsTab'    # Places panel in UTILITIES tab (workflow tools)


# Define your threaded insert configurations (name, diameter_mm, description):

THREADED_INSERT_CONFIGS = [
    ('M2_Insert', 3.200099, 'M3 threaded insert diameter'),
    ('M3_Insert', 4.250099, 'M4 threaded insert diameter'),
    # Add more insert sizes as needed: ('Parameter_Name', diameter_in_mm, 'Description')
]

# --------- END ----------------------------------------------------

# Generate derived data structures from the main configuration
THREADED_INSERT_PARAMETERS = [
    (name, f'{diameter} mm', description) 
    for name, diameter, description in THREADED_INSERT_CONFIGS
]

THREADED_INSERT_DIAMETERS = {
    f'{description.split()[0]} {description.split()[1]} {description.split()[2]}': diameter
    for name, diameter, description in THREADED_INSERT_CONFIGS
}

# Command IDs
ADD_PARAMS_CMD_ID = 'ThreadedInsertAddParams'
APPLY_APPEARANCE_CMD_ID = 'ThreadedInsertApplyAppearance'

def run(context):
    """Called when the add-in is loaded"""
    global app, ui
    try:
        app = adsk.core.Application.get()
        ui = app.userInterface
        
        # Create the command definitions
        create_add_parameters_command()
        create_apply_appearance_command()
        
        # Add buttons to the UI
        add_buttons_to_ui(context)
        
    except:
        if ui:
            ui.messageBox('Failed to initialize ThreadedInsertAddin:\n{}'.format(traceback.format_exc()))

def stop(context):
    """Called when the add-in is unloaded"""
    pass

def create_add_parameters_command():
    """Create the Add Parameters command definition"""
    global ui
    
    # Check if command already exists
    cmd_def = ui.commandDefinitions.itemById(ADD_PARAMS_CMD_ID)
    if cmd_def:
        cmd_def.deleteMe()
    
    # Create the command definition
    cmd_def = ui.commandDefinitions.addButtonDefinition(
        ADD_PARAMS_CMD_ID,
        'Add Insert Parameters',
        'Add threaded insert user parameters to the current design',
        './resources/AddParams'
    )
    
    # Connect to command created event
    on_command_created = AddParametersCommandCreatedHandler()
    cmd_def.commandCreated.add(on_command_created)
    handlers.append(on_command_created)

def create_apply_appearance_command():
    """Create the Apply Appearance command definition"""
    global ui
    
    # Check if command already exists
    cmd_def = ui.commandDefinitions.itemById(APPLY_APPEARANCE_CMD_ID)
    if cmd_def:
        cmd_def.deleteMe()
    
    # Create the command definition
    cmd_def = ui.commandDefinitions.addButtonDefinition(
        APPLY_APPEARANCE_CMD_ID,
        'Apply Insert Appearance',
        'Apply brass appearance to threaded insert holes',
        './resources/ApplyAppearance'
    )
    
    # Connect to command created event
    on_command_created = ApplyAppearanceCommandCreatedHandler()
    cmd_def.commandCreated.add(on_command_created)
    handlers.append(on_command_created)

def add_buttons_to_ui(context):
    """Add buttons to Fusion's UI"""
    global ui
    
    try:
        # Get the Design workspace
        design_workspace = ui.workspaces.itemById('FusionSolidEnvironment')
        if not design_workspace:
            ui.messageBox('Could not find Design workspace')
            return
        
        # Get the target tab (configured at top of file)
        design_tab = design_workspace.toolbarTabs.itemById(TARGET_TAB_ID)
        if not design_tab:
            ui.messageBox(f'Could not find tab with ID: {TARGET_TAB_ID}')
            return
        
        # Create our custom panel
        panel = design_tab.toolbarPanels.itemById('ThreadedInsertPanel')
        if not panel:
            panel = design_tab.toolbarPanels.add('ThreadedInsertPanel', 'Threaded Inserts', 'SelectPanel', False)
        
        # Add the buttons to the panel and promote them to be visible by default
        add_params_cmd = ui.commandDefinitions.itemById(ADD_PARAMS_CMD_ID)
        apply_appearance_cmd = ui.commandDefinitions.itemById(APPLY_APPEARANCE_CMD_ID)
        
        if add_params_cmd and not panel.controls.itemById(ADD_PARAMS_CMD_ID):
            add_params_control = panel.controls.addCommand(add_params_cmd)
            add_params_control.isPromotedByDefault = True
            add_params_control.isPromoted = True
        
        if apply_appearance_cmd and not panel.controls.itemById(APPLY_APPEARANCE_CMD_ID):
            apply_appearance_control = panel.controls.addCommand(apply_appearance_cmd)
            apply_appearance_control.isPromotedByDefault = True
            apply_appearance_control.isPromoted = True
        
    except Exception as e:
        ui.messageBox(f'Error adding buttons to UI: {str(e)}\n{traceback.format_exc()}')

# Event handlers for the commands
class AddParametersCommandCreatedHandler(adsk.core.CommandCreatedEventHandler):
    def __init__(self):
        super().__init__()
    
    def notify(self, args):
        try:
            cmd = args.command
            on_execute = AddParametersCommandExecuteHandler()
            cmd.execute.add(on_execute)
            handlers.append(on_execute)
        except:
            if ui:
                ui.messageBox('Failed in AddParametersCommandCreatedHandler:\n{}'.format(traceback.format_exc()))

class AddParametersCommandExecuteHandler(adsk.core.CommandEventHandler):
    def __init__(self):
        super().__init__()
    
    def notify(self, args):
        try:
            add_threaded_insert_parameters()
        except:
            if ui:
                ui.messageBox('Failed in AddParametersCommandExecuteHandler:\n{}'.format(traceback.format_exc()))

class ApplyAppearanceCommandCreatedHandler(adsk.core.CommandCreatedEventHandler):
    def __init__(self):
        super().__init__()
    
    def notify(self, args):
        try:
            cmd = args.command
            on_execute = ApplyAppearanceCommandExecuteHandler()
            cmd.execute.add(on_execute)
            handlers.append(on_execute)
        except:
            if ui:
                ui.messageBox('Failed in ApplyAppearanceCommandCreatedHandler:\n{}'.format(traceback.format_exc()))

class ApplyAppearanceCommandExecuteHandler(adsk.core.CommandEventHandler):
    def __init__(self):
        super().__init__()
    
    def notify(self, args):
        try:
            apply_threaded_insert_appearance()
        except:
            if ui:
                ui.messageBox('Failed in ApplyAppearanceCommandExecuteHandler:\n{}'.format(traceback.format_exc()))

# Core functionality functions (adapted from the scripts)
def add_threaded_insert_parameters():
    """Add threaded insert parameters to the current design"""
    global app, ui
    
    try:
        # Get the active design
        design = adsk.fusion.Design.cast(app.activeProduct)
        if not design:
            ui.messageBox('No active design found.')
            return
        
        # Get the user parameters collection
        user_parameters = design.userParameters
        
        # Check which parameters already exist and validate their values
        existing_parameters = {}
        for i in range(user_parameters.count):
            param = user_parameters.item(i)
            existing_parameters[param.name] = param.expression
        
        # Add new parameters and track results
        added_parameters = []
        skipped_parameters = []
        conflicting_parameters = []
        failed_parameters = []
        
        for param_name, param_value, param_comment in THREADED_INSERT_PARAMETERS:
            if param_name in existing_parameters:
                # Check if existing value matches expected value
                existing_value = existing_parameters[param_name]
                expected_value = param_value.replace(' mm', '')  # Remove unit for comparison
                
                # Compare values - handle different formats
                values_match = False
                try:
                    # Try numeric comparison first
                    existing_float = float(existing_value)
                    expected_float = float(expected_value)
                    values_match = abs(existing_float - expected_float) < 1e-8
                except:
                    # Fall back to string comparison
                    # Remove any units and whitespace for comparison
                    existing_clean = existing_value.replace('mm', '').replace(' ', '')
                    expected_clean = expected_value.replace('mm', '').replace(' ', '')
                    values_match = existing_clean == expected_clean
                
                if values_match:
                    skipped_parameters.append(f'{param_name} (already exists with correct value)')
                else:
                    conflicting_parameters.append(
                        f'{param_name} (existing: {existing_value}, expected: {expected_value}mm)'
                    )
                continue
            
            # Create ValueInput for the parameter
            value_input = adsk.core.ValueInput.createByString(param_value)
            
            # Add the user parameter with 'mm' unit
            new_parameter = user_parameters.add(param_name, value_input, 'mm', param_comment)
            if new_parameter:
                added_parameters.append(f'{param_name} = {param_value}')
            else:
                failed_parameters.append(f'{param_name} (creation failed)')
        
        # Build summary message
        summary_parts = []
        
        if added_parameters:
            summary_parts.append(f'✅ Added {len(added_parameters)} new parameters:')
            for added_param in added_parameters:
                summary_parts.append(f'  • {added_param}')
        
        if skipped_parameters:
            summary_parts.append(f'\n✅ Skipped {len(skipped_parameters)} parameters with correct values:')
            for skipped_param in skipped_parameters:
                summary_parts.append(f'  • {skipped_param}')
        
        if conflicting_parameters:
            summary_parts.append(f'\n⚠️ WARNING: {len(conflicting_parameters)} parameters have different values:')
            for conflicting_param in conflicting_parameters:
                summary_parts.append(f'  • {conflicting_param}')
            summary_parts.append('\n   Consider updating these manually to ensure proper threaded insert detection.')
        
        if failed_parameters:
            summary_parts.append(f'\n❌ Failed to add {len(failed_parameters)} parameters:')
            for failed_param in failed_parameters:
                summary_parts.append(f'  • {failed_param}')
        
        if not summary_parts:
            summary_parts.append('No changes made - all parameters already exist.')
        
        # Show results
        summary_text = '\n'.join(summary_parts)
        ui.messageBox(f'Threaded Insert Parameters Update:\n\n{summary_text}')
        
    except:
        if ui:
            ui.messageBox('Failed to add parameters:\n{}'.format(traceback.format_exc()))

def apply_threaded_insert_appearance():
    """Apply brass appearance to threaded insert holes"""
    global app, ui
    
    try:
        # Get the active design
        design = adsk.fusion.Design.cast(app.activeProduct)
        if not design:
            ui.messageBox('No active design found.')
            return
        
        # Find all threaded insert faces
        all_insert_faces = find_all_threaded_insert_faces(design)
        
        # Count total faces found
        total_faces = 0
        for faces in all_insert_faces.values():
            total_faces += len(faces)
        
        if total_faces == 0:
            diameter_list = ', '.join([f'{d}mm' for d in THREADED_INSERT_DIAMETERS.values()])
            ui.messageBox(f'No holes found with threaded insert diameters: {diameter_list}')
            return
        
        # Get or create the brass appearance
        brass_appearance = get_or_create_brass_appearance(design)
        if not brass_appearance:
            ui.messageBox('Failed to get or create "Brass - Matte" appearance.')
            return
        
        # Apply appearance to all matching cylindrical faces
        faces_processed = 0
        results_summary = []
        
        for insert_type, faces in all_insert_faces.items():
            if faces:
                type_count = 0
                for face in faces:
                    try:
                        face.appearance = brass_appearance
                        faces_processed += 1
                        type_count += 1
                    except Exception as e:
                        ui.messageBox(f'Failed to apply appearance to {insert_type} face: {str(e)}')
                
                if type_count > 0:
                    diameter_mm = THREADED_INSERT_DIAMETERS[insert_type]
                    results_summary.append(f'{insert_type} ({diameter_mm}mm): {type_count} face(s)')
        
        # Show summary of results
        if faces_processed > 0:
            summary_text = 'Successfully applied "Brass - Matte" appearance to:\n\n' + '\n'.join(results_summary)
            ui.messageBox(summary_text)
        else:
            ui.messageBox('No faces were successfully processed.')
        
    except:
        if ui:
            ui.messageBox('Failed to apply appearance:\n{}'.format(traceback.format_exc()))

# Helper functions (from the original scripts)
def get_or_create_brass_appearance(design):
    """Get the "Brass - Matte" appearance from the design or copy it from the library"""
    global app, ui
    
    # First, check if "Brass - Matte" already exists in the design
    existing_appearance = design.appearances.itemByName('Brass - Matte')
    if existing_appearance:
        return existing_appearance
    
    # Get the Fusion Appearance Library and copy "Brass - Matte"
    material_libraries = app.materialLibraries
    fusion_appearance_lib = material_libraries.itemByName('Fusion Appearance Library')
    
    if not fusion_appearance_lib:
        ui.messageBox('Could not find "Fusion Appearance Library"')
        return None
        
    brass_appearance = fusion_appearance_lib.appearances.itemByName('Brass - Matte')
    if not brass_appearance:
        ui.messageBox('Could not find "Brass - Matte" appearance in Fusion Appearance Library')
        return None
    
    # Copy the appearance to the design
    copied_appearance = design.appearances.addByCopy(brass_appearance, 'Brass - Matte')
    return copied_appearance

def find_all_threaded_insert_faces(design):
    """Find cylindrical faces that belong to holes with any threaded insert diameter"""
    all_matching_faces = {}
    tolerance = 1e-6  # Small tolerance for floating point comparison
    
    # Initialize results dictionary
    for insert_type in THREADED_INSERT_DIAMETERS.keys():
        all_matching_faces[insert_type] = []
    
    # Get all bodies in the design
    root_component = design.rootComponent
    all_bodies = []
    
    # Collect bodies from root component
    for i in range(root_component.bRepBodies.count):
        body = root_component.bRepBodies.item(i)
        all_bodies.append(body)
    
    # Collect bodies from all occurrences
    for i in range(root_component.allOccurrences.count):
        occurrence = root_component.allOccurrences.item(i)
        if occurrence.component:
            for j in range(occurrence.component.bRepBodies.count):
                body = occurrence.component.bRepBodies.item(j)
                all_bodies.append(body)
    
    # Check each body for matching cylindrical faces
    for body in all_bodies:
        for k in range(body.faces.count):
            face = body.faces.item(k)
            # Check against all threaded insert diameters
            for insert_type, diameter_mm in THREADED_INSERT_DIAMETERS.items():
                target_diameter_cm = diameter_mm / 10.0  # Convert mm to cm
                if is_matching_cylindrical_face(face, target_diameter_cm, tolerance):
                    all_matching_faces[insert_type].append(face)
                    break  # Don't check other diameters for this face
    
    return all_matching_faces

def is_matching_cylindrical_face(face, target_diameter, tolerance):
    """Check if a face is a cylindrical surface with the target diameter"""
    try:
        # Check if the face is cylindrical
        if face.geometry.surfaceType != adsk.core.SurfaceTypes.CylinderSurfaceType:
            return False
        
        # Cast to cylinder surface to get radius
        cylinder_surface = adsk.core.Cylinder.cast(face.geometry)
        if not cylinder_surface:
            return False
        
        # Get the radius and convert to diameter
        radius = cylinder_surface.radius
        diameter = radius * 2.0
        
        # Check if diameter matches target within tolerance
        diameter_diff = abs(diameter - target_diameter)
        
        # Additional check: ensure this is likely an internal face (hole)
        is_internal = is_internal_cylindrical_face(face)
        
        return diameter_diff <= tolerance and is_internal
        
    except:
        return False

def is_internal_cylindrical_face(face):
    """Determine if a cylindrical face is internal (like the inside of a hole)"""
    try:
        # Get a point on the face and evaluate the normal
        evaluator = face.evaluator
        if not evaluator:
            return False
        
        # Get the parameter range
        success, u_min, u_max, v_min, v_max = evaluator.parametricRange()
        if not success:
            return False
        
        # Evaluate at the middle of the parameter range
        u_mid = (u_min + u_max) / 2.0
        v_mid = (v_min + v_max) / 2.0
        
        success, point, normal = evaluator.getPointAtParameter(u_mid, v_mid)
        if not success:
            return False
        
        # For an internal cylindrical face (hole), the normal should point inward
        cylinder_surface = adsk.core.Cylinder.cast(face.geometry)
        if not cylinder_surface:
            return False
        
        # Vector from point on surface to cylinder axis
        axis_origin = cylinder_surface.origin
        to_axis = axis_origin.vectorTo(point)
        
        # For an internal face, the normal should be roughly opposite to the vector toward the axis
        dot_product = normal.dotProduct(to_axis)
        
        # If dot product is negative, normal points toward axis (internal face)
        return dot_product < 0
        
    except:
        # If we can't determine, default to True (assume internal)
        return True