"""Fusion 360 script entrypoint: list all components in the active design."""

from __future__ import annotations

import csv
import json
import os
from datetime import datetime
import traceback
import adsk.core
import adsk.fusion

# Initialize the global variables for the Application and UserInterface objects.
app = adsk.core.Application.get()
ui = app.userInterface

_handlers: list[adsk.core.EventHandler] = []
_last_csv_dir: str | None = None
_config_cache: dict[str, object] | None = None


def _get_active_design() -> adsk.fusion.Design:
    product = app.activeProduct
    design = adsk.fusion.Design.cast(product)
    if not design:
        app.log("BomSzlendakv2: No active design found.")
        raise RuntimeError("No active Fusion design. Open a design and try again.")
    return design


def _build_table(rows: list[tuple[str, str, int, str, str, str]]) -> str:
    header = "No. | Component Name | Material | Qty | X | Y | Z"
    separator = "-" * len(header)
    lines = [header, separator]
    for index, (name, material, qty, x_size, y_size, z_size) in enumerate(rows, start=1):
        lines.append(f"{index:>3} | {name} | {material} | {qty} | {x_size} | {y_size} | {z_size}")
    return "\n".join(lines)


def _get_component_size(
    component: adsk.fusion.Component, units_manager: adsk.core.UnitsManager
) -> tuple[str, str, str]:
    bodies = component.bRepBodies
    if bodies.count == 0:
        return ("-", "-", "-")

    min_x = min_y = min_z = float("inf")
    max_x = max_y = max_z = float("-inf")

    for index in range(bodies.count):
        bbox = bodies.item(index).boundingBox
        min_point = bbox.minPoint
        max_point = bbox.maxPoint
        min_x = min(min_x, min_point.x)
        min_y = min(min_y, min_point.y)
        min_z = min(min_z, min_point.z)
        max_x = max(max_x, max_point.x)
        max_y = max(max_y, max_point.y)
        max_z = max(max_z, max_point.z)

    size_x = max_x - min_x
    size_y = max_y - min_y
    size_z = max_z - min_z

    from_units = units_manager.internalUnits
    to_units = units_manager.defaultLengthUnits
    size_x = units_manager.convert(size_x, from_units, to_units)
    size_y = units_manager.convert(size_y, from_units, to_units)
    size_z = units_manager.convert(size_z, from_units, to_units)

    return (
        f"{size_x:.2f}",
        f"{size_y:.2f}",
        f"{size_z:.2f}",
    )




def _get_component_material(component: adsk.fusion.Component) -> str:
    try:
        material = component.material
        if material:
            return material.name
    except Exception:
        pass

    try:
        bodies = component.bRepBodies
        for index in range(bodies.count):
            body = bodies.item(index)
            material = body.material
            if material:
                return material.name
    except Exception:
        pass

    return ""

def _reorder_sizes_for_thickness(
    size_x: float, size_y: float, size_z: float, thickness_override: float | None
) -> tuple[float, float, float]:
    if thickness_override is not None:
        return size_x, size_y, thickness_override

    sizes = [("x", size_x), ("y", size_y), ("z", size_z)]
    sizes.sort(key=lambda item: item[1])
    thickness = sizes[0][1]
    remaining = [item[1] for item in sizes[1:]]
    return remaining[0], remaining[1], thickness


def _get_length_parameters(
    design: adsk.fusion.Design, units_manager: adsk.core.UnitsManager
) -> dict[str, adsk.fusion.Parameter]:
    params: dict[str, adsk.fusion.Parameter] = {}
    user_params = design.userParameters
    for index in range(user_params.count):
        param = user_params.item(index)
        try:
            if param.unitType == adsk.fusion.ParameterUnitTypes.LengthUnitType:
                params[param.name] = param
                continue
        except Exception:
            pass

        try:
            if units_manager.isValidExpression(
                param.expression, units_manager.defaultLengthUnits
            ):
                params[param.name] = param
        except Exception:
            continue
    return params


def _evaluate_parameter_length(
    param: adsk.fusion.Parameter, units_manager: adsk.core.UnitsManager
) -> float:
    target_units = units_manager.defaultLengthUnits
    value = units_manager.convert(param.value, units_manager.internalUnits, target_units)
    app.log(
        f"BomSzlendakv2: Thickness parameter '{param.name}' value "
        f"{value:.4f} {target_units}."
    )
    return value


def _show_table(
    design: adsk.fusion.Design,
    thickness_param: adsk.fusion.Parameter | None,
    csv_path: str | None,
    merge_enabled: bool,
) -> None:
    app.log("BomSzlendakv2: Building component table.")
    components = design.allComponents
    root_component = design.rootComponent
    units_manager = design.unitsManager

    thickness_override = None
    if thickness_param is not None:
        app.log(f"BomSzlendakv2: Using thickness parameter '{thickness_param.name}'.")
        thickness_override = _evaluate_parameter_length(thickness_param, units_manager)
    else:
        app.log("BomSzlendakv2: Using auto thickness (smallest dimension).")

    counts: dict[str, int] = {}
    occurrences = root_component.allOccurrences
    for index in range(occurrences.count):
        occurrence = occurrences.item(index)
        component = occurrence.component
        token = component.entityToken
        counts[token] = counts.get(token, 0) + 1

    rows: list[tuple[str, str, int, str, str, str]] = []
    for index in range(components.count):
        component = components.item(index)
        token = component.entityToken
        qty = counts.get(token, 0)
        if component == root_component:
            qty = max(qty, 1)

        material = _get_component_material(component)

        size_x, size_y, size_z = _get_component_size(component, units_manager)
        if size_x != "-" and thickness_override is not None:
            x_val = float(size_x)
            y_val = float(size_y)
            z_val = float(size_z)
        elif size_x != "-":
            x_val = float(size_x)
            y_val = float(size_y)
            z_val = float(size_z)
        else:
            x_val = y_val = z_val = 0.0

        if size_x == "-":
            rows.append((component.name, material, qty, "-", "-", "-"))
        else:
            x_val, y_val, z_val = _reorder_sizes_for_thickness(
                x_val, y_val, z_val, thickness_override
            )
            rows.append(
                (
                    component.name,
                    material,
                    qty,
                    f"{x_val:.2f}",
                    f"{y_val:.2f}",
                    f"{z_val:.2f}",
                )
            )

    if merge_enabled:
        rows = _merge_rows(rows)

    rows.sort(key=lambda value: (value[1] == "", value[1].casefold(), value[0].casefold()))

    units_label = units_manager.defaultLengthUnits
    header = f"BOM list by Szlendak (units: {units_label})"
    table = _build_table(rows) if rows else "No components found in the design."

    project_name = _get_project_name()
    export_dt = _format_export_datetime()

    if csv_path:
        csv_path = _export_csv(rows, units_label, project_name, export_dt, csv_path)
        ui.messageBox(
            f"{header}\n\n{table}\n\nCSV saved to:\n{csv_path}", "Components in Design"
        )
    else:
        ui.messageBox(f"{header}\n\n{table}", "Components in Design")


def _merge_rows(
    rows: list[tuple[str, str, int, str, str, str]]
) -> list[tuple[str, str, int, str, str, str]]:
    merged_rows: list[tuple[str, str, int, str, str, str]] = []
    index_by_key: dict[tuple[str, str, str, str], int] = {}

    for name, material, qty, x_size, y_size, z_size in rows:
        if x_size == "-" or y_size == "-" or z_size == "-":
            merged_rows.append((name, material, qty, x_size, y_size, z_size))
            continue

        key = (material, x_size, y_size, z_size)
        existing_index = index_by_key.get(key)
        if existing_index is None:
            index_by_key[key] = len(merged_rows)
            merged_rows.append((name, material, qty, x_size, y_size, z_size))
            continue

        existing = merged_rows[existing_index]
        existing_name, existing_material, existing_qty, ex_x, ex_y, ex_z = existing
        new_qty = existing_qty + qty
        if "(merged)" not in existing_name:
            existing_name = f"{existing_name} (merged)"
        merged_rows[existing_index] = (
            existing_name,
            existing_material,
            new_qty,
            ex_x,
            ex_y,
            ex_z,
        )

    return merged_rows


def _format_export_datetime() -> str:
    # "Pp" style: long date + short time
    return datetime.now().strftime("%B %d, %Y %H:%M")


def _get_project_name() -> str:
    try:
        document = app.activeDocument
        if document:
            return document.name
    except Exception:
        pass
    return ""


def _get_default_downloads_path() -> str:
    home = os.path.expanduser("~")
    downloads = os.path.join(home, "Downloads")
    if os.path.isdir(downloads):
        return downloads
    return os.path.dirname(os.path.abspath(__file__))


def _get_default_csv_dir() -> str:
    if _last_csv_dir and os.path.isdir(_last_csv_dir):
        return _last_csv_dir
    saved = _get_saved_csv_dir()
    if saved:
        return saved
    return _get_default_downloads_path()


def _sanitize_filename(value: str) -> str:
    if not value:
        return "Project"
    cleaned = "".join(ch if ch.isalnum() or ch in (" ", "-", "_") else "_" for ch in value)
    cleaned = cleaned.strip().replace("  ", " ")
    cleaned = cleaned.replace(" ", "_")
    return cleaned or "Project"


def _default_csv_filename(project_name: str) -> str:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_name = _sanitize_filename(project_name)
    return f"{safe_name}_BOM_{timestamp}.csv"


def _show_save_dialog(initial_dir: str, suggested_name: str) -> str | None:
    file_dialog = ui.createFileDialog()
    file_dialog.title = "Save BOM CSV"
    file_dialog.filter = "CSV Files (*.csv)"
    file_dialog.filterIndex = 0
    file_dialog.initialDirectory = initial_dir
    if hasattr(file_dialog, "initialFilename"):
        file_dialog.initialFilename = suggested_name
    elif hasattr(file_dialog, "initialFileName"):
        file_dialog.initialFileName = suggested_name

    if file_dialog.showSave() == adsk.core.DialogResults.DialogOK:
        return file_dialog.filename
    return None


def _remember_csv_dir(path: str) -> None:
    global _last_csv_dir
    if not path:
        return
    directory = os.path.dirname(path)
    if directory:
        _last_csv_dir = directory
        _save_config({"last_csv_dir": _last_csv_dir})


def _config_path() -> str:
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")


def _load_config() -> dict[str, object]:
    global _config_cache
    if _config_cache is not None:
        return _config_cache
    path = _config_path()
    if not os.path.isfile(path):
        _config_cache = {}
        return _config_cache
    try:
        with open(path, "r", encoding="utf-8") as handle:
            _config_cache = json.load(handle)
            if not isinstance(_config_cache, dict):
                _config_cache = {}
    except Exception:
        _config_cache = {}
    return _config_cache


def _save_config(update: dict[str, object]) -> None:
    data = _load_config()
    data.update(update)
    path = _config_path()
    try:
        with open(path, "w", encoding="utf-8") as handle:
            json.dump(data, handle, ensure_ascii=True, indent=2)
    except Exception:
        app.log(f"Failed:\n{traceback.format_exc()}")


def _get_saved_csv_dir() -> str | None:
    value = _load_config().get("last_csv_dir")
    if isinstance(value, str) and os.path.isdir(value):
        return value
    return None


def _get_saved_export_enabled() -> bool:
    value = _load_config().get("export_csv_enabled")
    if isinstance(value, bool):
        return value
    return True


def _get_saved_merge_enabled() -> bool:
    value = _load_config().get("merge_same_size_enabled")
    if isinstance(value, bool):
        return value
    return False


def _export_csv(
    rows: list[tuple[str, str, int, str, str, str]],
    units_label: str,
    project_name: str,
    export_dt: str,
    path: str,
) -> str:
    app.log(f"BomSzlendakv2: Exporting CSV to {path}.")

    with open(path, "w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(
            [
                f"BOM list by Szlendak (units: {units_label})",
                project_name,
                export_dt,
            ]
        )
        writer.writerow([])
        writer.writerow(["No.", "Component Name", "Material", "Qty", "X", "Y", "Z"])
        for index, (name, material, qty, x_size, y_size, z_size) in enumerate(rows, start=1):
            writer.writerow([index, name, material, qty, x_size, y_size, z_size])

    return path


def run(_context: str):
    """This function is called by Fusion when the script is run."""
    try:
        app.log("BomSzlendakv2: run() started.")
        adsk.autoTerminate(False)
        design = _get_active_design()
        units_manager = design.unitsManager
        length_params = _get_length_parameters(design, units_manager)
        app.log(f"BomSzlendakv2: Found {len(length_params)} length parameters.")

        cmd_def_id = "BomSzlendakv2_ListComponents"
        existing_cmd_def = ui.commandDefinitions.itemById(cmd_def_id)
        if existing_cmd_def:
            app.log("BomSzlendakv2: Removing existing command definition.")
            existing_cmd_def.deleteMe()

        cmd_def = ui.commandDefinitions.addButtonDefinition(
            cmd_def_id,
            "Szlendak BOM List",
            "List all components with quantity and sizes.",
        )
        app.log("BomSzlendakv2: Command definition created.")

        class CommandCreatedHandler(adsk.core.CommandCreatedEventHandler):
            def __init__(self):
                super().__init__()

            def notify(self, args: adsk.core.CommandCreatedEventArgs):
                try:
                    app.log("BomSzlendakv2: CommandCreated fired.")
                    command = args.command
                    inputs = command.commandInputs

                    dropdown = inputs.addDropDownCommandInput(
                        "thicknessParam",
                        "Thickness parameter (optional)",
                        adsk.core.DropDownStyles.TextListDropDownStyle,
                    )
                    dropdown.listItems.add("None (auto thickness)", True, "")
                    for name in sorted(length_params.keys(), key=str.casefold):
                        dropdown.listItems.add(name, False, "")
                    app.log(
                        "BomSzlendakv2: Dropdown populated with length parameters."
                    )

                    project_name = _get_project_name()
                    default_name = _default_csv_filename(project_name)
                    export_enabled = _get_saved_export_enabled()
                    export_checkbox = inputs.addBoolValueInput(
                        "exportCsv",
                        "Export CSV",
                        True,
                        "",
                        export_enabled,
                    )
                    csv_path_input = inputs.addStringValueInput(
                        "csvPath",
                        "CSV output path",
                        os.path.join(_get_default_csv_dir(), default_name),
                    )
                    browse_input = inputs.addBoolValueInput(
                        "browseCsv",
                        "Browse...",
                        False,
                        "",
                        False,
                    )
                    merge_enabled = _get_saved_merge_enabled()
                    merge_checkbox = inputs.addBoolValueInput(
                        "mergeSameSize",
                        "Merge same size",
                        True,
                        "",
                        merge_enabled,
                    )
                    csv_path_input.isEnabled = export_checkbox.value
                    browse_input.isEnabled = export_checkbox.value

                    class InputChangedHandler(adsk.core.InputChangedEventHandler):
                        def __init__(self):
                            super().__init__()

                        def notify(self, input_args: adsk.core.InputChangedEventArgs):
                            try:
                                changed = input_args.input
                                if changed.id == "exportCsv":
                                    csv_path_input.isEnabled = export_checkbox.value
                                    browse_input.isEnabled = export_checkbox.value
                                    _save_config(
                                        {"export_csv_enabled": export_checkbox.value}
                                    )
                                    return
                                if changed.id == "mergeSameSize":
                                    _save_config(
                                        {"merge_same_size_enabled": merge_checkbox.value}
                                    )
                                    return
                                if changed.id != "browseCsv":
                                    return

                                app.log("BomSzlendakv2: Browse CSV clicked.")
                                path = _show_save_dialog(
                                    _get_default_csv_dir(),
                                    default_name,
                                )
                                if path:
                                    csv_path_input.value = path
                                    _remember_csv_dir(path)
                                changed.value = False
                            except Exception:
                                app.log(f"Failed:\n{traceback.format_exc()}")

                    class ExecuteHandler(adsk.core.CommandEventHandler):
                        def __init__(self):
                            super().__init__()

                        def notify(self, execute_args: adsk.core.CommandEventArgs):
                            try:
                                app.log("BomSzlendakv2: Execute fired.")
                                selected = dropdown.selectedItem.name
                                param = length_params.get(selected)
                                csv_path = None
                                if export_checkbox.value:
                                    csv_path = csv_path_input.value.strip()
                                    if not csv_path:
                                        csv_path = _show_save_dialog(
                                            _get_default_csv_dir(),
                                            default_name,
                                        )
                                    if csv_path:
                                        _remember_csv_dir(csv_path)
                                _show_table(design, param, csv_path, merge_checkbox.value)
                                adsk.autoTerminate(True)
                            except Exception:
                                app.log(f"Failed:\n{traceback.format_exc()}")
                                ui.messageBox(
                                    "Failed to list components. "
                                    "See the Text Commands window for details."
                                )
                                adsk.autoTerminate(True)

                    input_handler = InputChangedHandler()
                    command.inputChanged.add(input_handler)
                    _handlers.append(input_handler)

                    exec_handler = ExecuteHandler()
                    command.execute.add(exec_handler)
                    _handlers.append(exec_handler)
                except Exception:
                    app.log(f"Failed:\n{traceback.format_exc()}")

        created_handler = CommandCreatedHandler()
        cmd_def.commandCreated.add(created_handler)
        _handlers.append(created_handler)

        app.log("BomSzlendakv2: Executing command definition.")
        cmd_def.execute()
        app.log("BomSzlendakv2: run() finished.")
    except Exception:  # pylint: disable=broad-except
        app.log(f"Failed:\n{traceback.format_exc()}")
        ui.messageBox("Failed to list components. See the Text Commands window for details.")
