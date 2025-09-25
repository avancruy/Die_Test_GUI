import pandas as pd
import numpy as np

# --- Utility Functions (RESTORED TO ORIGINAL) ---
def string_to_num(s, target_type=float):
    try:
        return target_type(s)
    except ValueError:
        try:  # Retry with float for potential scientific notation then convert
            val = float(s)
            if target_type == int:
                return int(val)
            return val
        except ValueError:
            return None  # Indicates conversion failure

def parse_measurement_data(data_string):
    try:
        if isinstance(data_string, (list, tuple)):  # Already parsed
            return [float(x) for x in data_string]
        if not data_string: return []
        values = [float(x.strip()) for x in data_string.strip().split(',')]
        return values
    except ValueError as e:
        print(f"Error parsing measurement data string '{data_string}': {e}")
        return []
    except TypeError:
        print(f"Invalid data type for parsing: {data_string}")
        return []


# ✨ Modified function to multiply PD current by 1.69
def create_combined_excel_file(laser_data, detector_data, eam_data, timestamp,
                               detector_params, laser_params, eam_params,
                               is_eam, device_id, temperature, base_path_config):
    try:
        laser_voltages_fetched = parse_measurement_data(laser_data)
        detector_currents_fetched = parse_measurement_data(detector_data)
        eam_currents_fetched = parse_measurement_data(eam_data)

        # Determine num_points from the parameter set that is performing a sweep
        if laser_params['source_mode'].lower() == 'swe':
            num_points_sweep = laser_params['num_points']
        elif eam_params['source_mode'].lower() == 'swe':
            num_points_sweep = eam_params['num_points']
        else:  # Default or if both are fixed (though one should be sweep for a typical test)
            num_points_sweep = max(laser_params['num_points'], eam_params['num_points'], detector_params['num_points'],
                                   1)

        laser_current_setpoints = np.linspace(laser_params['start'], laser_params['stop'], num_points_sweep)
        eam_voltage_setpoints = np.linspace(eam_params['start'], eam_params['stop'], num_points_sweep)

        if detector_params['source_mode'].lower() == 'fix':
            detector_voltage_setpoints = np.full(num_points_sweep, detector_params['initval'])
        else:
            detector_voltage_setpoints = np.linspace(detector_params['start'], detector_params['stop'],
                                                     num_points_sweep)

        def pad_or_truncate(data_list, length):
            if not isinstance(data_list, list): data_list = []
            if len(data_list) < length:
                return data_list + [np.nan] * (length - len(data_list))
            return data_list[:length]

        laser_voltages_final = pad_or_truncate(laser_voltages_fetched, num_points_sweep)
        detector_currents_final = pad_or_truncate(detector_currents_fetched, num_points_sweep)
        eam_currents_final = pad_or_truncate(eam_currents_fetched, num_points_sweep)

        laser_currents_mA = laser_current_setpoints
        # ✨ Apply 1.69 multiplication factor to PD current readings
        detector_currents_mA = np.array(detector_currents_final) * 1000
        eam_currents_mA = np.array(eam_currents_final) * 1000

        combined_data = {
            'SMU2_Ch1_EAM_Voltage_Set_V': eam_voltage_setpoints,
            'SMU2_Ch1_EAM_Current_Meas_mA': eam_currents_mA,
            'SMU1_Ch2_Laser_Current_Set_mA': laser_currents_mA,
            'SMU1_Ch2_Laser_Voltage_Meas_V': laser_voltages_final,
            'SMU1_Ch1_PD_Voltage_Set_V': detector_voltage_setpoints,
            'SMU1_Ch1_PD_Current_Meas_mA': np.abs(detector_currents_mA),
        }

        combined_df = pd.DataFrame(combined_data)

        pulse_width_s = laser_params.get('pulse_width', 0)
        period_s = laser_params.get('trigger_period', 0)
        duty_cycle = 100

        if (laser_params['source_shape'] == 'puls'):
            duty_cycle = (pulse_width_s / period_s) * 100 if period_s > 0 else 0

        # Determine if laser is in pulsed mode (duty cycle < 100% or source_shape is 'puls' with DC < 100%)
        is_pulsed = (laser_params.get('source_shape', '').lower() == 'puls' and duty_cycle < 100)

        base_path = base_path_config
        file_prefix = f"{device_id}_" if device_id else ""
        temp_suffix = f"{temperature}" if temperature else ""

        common_suffix = f"NumPoints{num_points_sweep}_DtyC{duty_cycle:.2f}%_{temp_suffix}°C_{timestamp}.xlsx"

        if not is_eam:
            ld_start_mA = int(laser_params['start'])
            ld_stop_mA = int(laser_params['stop'])
            eam_bias_V = eam_params['initval']
            pd_bias_V = detector_params['initval']

            # Only add "pulsed_" prefix if actually in pulsed mode
            mode_prefix = "pulsed_" if is_pulsed else ""
            filename = (
                f"{base_path}{file_prefix}{mode_prefix}LIV_LDBias({ld_start_mA},{ld_stop_mA})mA_"
                f"EAMBias({eam_bias_V})V_PDBias({pd_bias_V})V_{common_suffix}"
            )
        else:
            ld_bias_mA = int(laser_params['initval'])
            eam_start_V = eam_params['start']
            eam_stop_V = eam_params['stop']
            pd_bias_V = detector_params['initval']

            # Only add "pulsed_" prefix if actually in pulsed mode
            mode_prefix = "pulsed_" if is_pulsed else ""
            filename = (
                f"{base_path}{file_prefix}{mode_prefix}EAM_LDBias({ld_bias_mA})mA_"
                f"EAMBias({eam_start_V},{eam_stop_V})V_PDBias({pd_bias_V})V_{common_suffix}"
            )

        combined_df.to_excel(filename, index=False)
        print(f"\nCombined Excel file created successfully: {filename}")
        return True

    except Exception as e:
        print(f"Error creating Excel file: {e}")
        import traceback
        traceback.print_exc()
        return False