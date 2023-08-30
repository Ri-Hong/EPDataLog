#include <windows.h>
#include <stdio.h>
#include <time.h>

#define READ_BAT_TOTAL_VOLTAGE_CURRENT_SOC		    0x90
#define READ_BAT_HIGHEST_LOWEST_VOLTAGE		        0x91
#define READ_BAT_MAX_MIN_TEMP		                0x92
#define READ_BAT_CHARGE_DISCHARGE_MOS_STATUS		0x93
#define READ_BAT_STATUS_INFO_1		                0x94
#define READ_BAT_SINGLE_CELL_VOLTAGE		        0x95
#define READ_BAT_SINGLE_CELL_TEMP		            0x96
#define READ_BAT_SINGLE_CELL_BALANCE_STATUS		    0x97 
#define READ_BAT_SINGLE_CELL_FAILURE_STATUS		    0x98


#define G_MAX_NUMBER_OF_CELLS                       16
#define G_MAX_NUMBER_OF_TEMP_SENSORS                4

// Global Vars
int g_number_of_battery_cells = -1;
int g_number_of_temp_sensors = -1;
int g_delay_time_ms = 2000;
const int MIN_DELAY_TIME = 0;
const int MAX_COM_PORT_NUMBER = 256;
const int MIN_COM_PORT_NUMBER = 1;

const int INVALID_COM_PORT_NUMBER = -2;
const int INVALID_DELAY_TIME_SUPPLIED = -3;
const int NO_COM_PORT_NUMBER_SUPPLIED = -1;
const int NO_DELAY_TIME_SUPPLIED = -1;

int readProgramParams(int argc, char *argv[]);
int isInteger(char *str);
HANDLE setupCOMPort(int portNumber);
HANDLE connectToCOMPort(const char *portName);
int getDateTime();
int getBMSData(HANDLE hComm, int requestType);
int parseBmsResponseSoc(unsigned char *pResponse);
int parseBmsResponseHighestLowestVoltage(unsigned char *pResponse);
int parseBmsResponseMaxMinTemp(unsigned char *pResponse);
int parseBmsResponseChargeDischargeMosStatus(unsigned char *pResponse);
int parseBmsResponseStatusInfo1(unsigned char *pResponse);
int parseBmsResponseSingleCellVoltage(unsigned char *pResponse);
int parseBmsResponseSingleCellTemp(unsigned char *pResponse);
int parseBmsResponseSingleCellBalancingStatus(unsigned char *pResponse);
int parseBmsResponseBatteryFailureStatus(unsigned char *pResponse);
FILE *openCsvFile();
int printCsvHeader(FILE *fp);
int outputBMSDataToCsv(FILE *fp);

// BMS Data Structure
typedef struct {
    int lineNumber;
    char dateTime[20];  // Enough to hold "YYYY-MM-DD HH:MM:SS" and '\0'
    int batteryID;
    float current;
    float voltage;
    float stateOfCharge;
    float totalCapacity;
    float remainingCapacity;
    float cellVoltage[G_MAX_NUMBER_OF_CELLS];
    float highestCellVoltage;
    float lowestCellVoltage;
    float temperatures[G_MAX_NUMBER_OF_TEMP_SENSORS];
    int chargingDischargingStatus;
    int chargingMOSStatus;
    int dischargingMOSStatus;
    int balancingStatus;
    // Balance Status' are stored in reverse order. I.e. balance_status[0] is the balance status of the last cell
    int cellBalancingStatus[G_MAX_NUMBER_OF_CELLS];
    // Alarms' are stored in reverse bit order. I.e. alarms[i][0] is the alarm of the ith byte, last bit
    char alarms[8][9];

} BMSData;

BMSData bmsData; // Declare a BMSData struct variable


// Returns the supplied COM port number, or -1 if none was supplied
int readProgramParams(int argc, char *argv[]) {
    int suppliedComPortNumber = NO_COM_PORT_NUMBER_SUPPLIED;
    int suppliedDelayTime = NO_DELAY_TIME_SUPPLIED;

    // Try to read the COM port number from the command line
    for (int i = 1; i < argc; i++) {  // Starting at 1 to skip the program name itself
        if (strcmp(argv[i], "-c") == 0) {
            if (i + 1 < argc) {  // Make sure we don't go out of bounds
                if (!isInteger(argv[i + 1])) {
                    printf("Error: Invalid value for -c option. Aborting\n");
                    return INVALID_COM_PORT_NUMBER;
                }
                suppliedComPortNumber = atoi(argv[++i]);  // Convert next argument to integer and assign to t_value
                // Handle Error checking for com port number time
                if (suppliedComPortNumber < MIN_COM_PORT_NUMBER) {
                    printf("Error: COM port number cannot be less than %d. Aborting\n", MIN_COM_PORT_NUMBER);
                    suppliedComPortNumber = INVALID_COM_PORT_NUMBER;
                } else if (suppliedComPortNumber > MAX_COM_PORT_NUMBER) {
                    printf("Error: COM port number cannot be greater than %d. Aborting\n", MAX_COM_PORT_NUMBER);
                    suppliedComPortNumber = INVALID_COM_PORT_NUMBER;
                }
            } else {
                printf("Error: Missing value for -c option\n");
            }
    // Try to read the delay time from the command line
        } else if (strcmp(argv[i], "-t") == 0) {
            if (i + 1 < argc) {  // Make sure we don't go out of bounds
                if (!isInteger(argv[i + 1])) {
                    printf("Error: Invalid value for -t option\n");
                    continue;
                }
                suppliedDelayTime = atoi(argv[++i]);  // Convert next argument to integer and assign to c_value
                // Handle Error checking for delay time
                if (suppliedDelayTime < MIN_DELAY_TIME) {
                    printf("Error: Delay time cannot be less than %d. Aborting\n", MIN_DELAY_TIME);
                    return INVALID_DELAY_TIME_SUPPLIED;
                } else {
                    g_delay_time_ms = suppliedDelayTime;
                    printf("Success: Delay time of %dms set\n", g_delay_time_ms);
                }

            } else {
                printf("Error: Missing value for -t option\n");
            }
        }
    }

    if (suppliedDelayTime == NO_DELAY_TIME_SUPPLIED) {
        printf("No delay time supplied. Using default delay time of %dms\n", g_delay_time_ms);
    }

    return suppliedComPortNumber;

}

int isInteger(char *str) {
    char *endptr;
    strtol(str, &endptr, 10);
    if (endptr == str) {
        return 0;  // No conversion performed
    }
    if (*endptr != '\0') {
        return 0;  // String is partially a valid integer
    }
    return 1;
}

HANDLE setupCOMPort(int portNumber) {
    if (portNumber != -1) {
        char comPortName[10];  // Enough space for "COM" + up to 3 digits + null terminator
        sprintf(comPortName, "COM%d", portNumber);  // Format the COM port name
        printf("Attempting to use COM port %d\n", portNumber);
        HANDLE hComm = connectToCOMPort(comPortName);
        Sleep(500);
        if (hComm != INVALID_HANDLE_VALUE) {
            // Use REQUEST_TOTAL_VOLTAGE_CURRENT_SOC as queryData
            const unsigned char queryData[13] = {0XA5, 0X40, 0X90, 0X08, 0X00, 0X00, 0x00, 0x00, 0X00, 0X00, 0X00, 0X00, 0x7D};
            DWORD bytesWritten;

            // Send a query
            if (!WriteFile(hComm, queryData, sizeof(queryData), &bytesWritten, NULL)) {
                printf("Error in writing to COM port. Aborting.\n");
                CloseHandle(hComm);
                return INVALID_HANDLE_VALUE;
            }

            unsigned char buffer[300];
            DWORD bytesRead;

            // Read the response
            if (!ReadFile(hComm, buffer, sizeof(buffer), &bytesRead, NULL)) {
                printf("Error in reading from COM port. Aborting.\n");
                CloseHandle(hComm);
                return INVALID_HANDLE_VALUE;
            }

            // Check the response
            if (buffer[0] == 0xA5 && buffer[1] == 0x01 && buffer[2] == 0x90) {
                printf("Found the target COM port: %s\n", comPortName);
                return hComm;
            }

            // Close the handle if this isn't the correct one
            CloseHandle(hComm);
        } else {
            printf("Error: Unable to open COM port %s. Aborting.\n", comPortName);
            return INVALID_HANDLE_VALUE;
        }
    } else {
        printf("No COM port supplied. Searching for a COM port... \n");
        for (int i = 1; i <= 256; ++i) {
            char comPortName[10];  // Enough space for "COM" + up to 3 digits + null terminator
            sprintf(comPortName, "COM%d", i);  // Format the COM port name
            HANDLE hComm = connectToCOMPort(comPortName);
            Sleep(500);
            if (hComm != INVALID_HANDLE_VALUE) {
                // Use REQUEST_TOTAL_VOLTAGE_CURRENT_SOC as queryData
                const unsigned char queryData[13] = {0XA5, 0X40, 0X90, 0X08, 0X00, 0X00, 0x00, 0x00, 0X00, 0X00, 0X00, 0X00, 0x7D};
                DWORD bytesWritten;

                // Send a query
                if (!WriteFile(hComm, queryData, sizeof(queryData), &bytesWritten, NULL)) {
                    printf("Error in writing to COM port\n");
                    CloseHandle(hComm);
                    continue;
                }

                unsigned char buffer[300];
                DWORD bytesRead;

                // Read the response
                if (!ReadFile(hComm, buffer, sizeof(buffer), &bytesRead, NULL)) {
                    printf("Error in reading from COM port\n");
                    CloseHandle(hComm);
                    continue;
                }

                // Check the response
                if (buffer[0] == 0xA5 && buffer[1] == 0x01 && buffer[2] == 0x90) {
                    printf("Found the target COM port: %s\n", comPortName);
                    return hComm;
                }

                // Close the handle if this isn't the correct one
                CloseHandle(hComm);
            } else {
                printf("Unable to open COM port %s\n", comPortName);
            }
        }
    
    }
}

HANDLE connectToCOMPort(const char *portName) {
    printf("Trying port %s\n", portName);
    DCB dcbSerialParams = {0};
    COMMTIMEOUTS timeouts = {0};
    
    HANDLE hComm = CreateFile(portName,
                              GENERIC_READ | GENERIC_WRITE,
                              0,
                              NULL,
                              OPEN_EXISTING,
                              0,
                              NULL);

    if (hComm == INVALID_HANDLE_VALUE) {
        return INVALID_HANDLE_VALUE;
    }

    dcbSerialParams.DCBlength = sizeof(DCB);

    if (!GetCommState(hComm, &dcbSerialParams)) {
        printf("Error getting current DCB settings\n");
        CloseHandle(hComm);
        return INVALID_HANDLE_VALUE;
    }

    dcbSerialParams.BaudRate = CBR_9600;
    dcbSerialParams.ByteSize = 8;
    dcbSerialParams.StopBits = ONESTOPBIT;
    dcbSerialParams.Parity   = NOPARITY;

    if (!SetCommState(hComm, &dcbSerialParams)) {
        printf("Could not set serial port parameters\n");
        CloseHandle(hComm);
        return INVALID_HANDLE_VALUE;
    }

    timeouts.ReadIntervalTimeout         = 50;
    timeouts.ReadTotalTimeoutConstant    = 50;
    timeouts.ReadTotalTimeoutMultiplier  = 10;
    timeouts.WriteTotalTimeoutConstant   = 50;
    timeouts.WriteTotalTimeoutMultiplier = 10;

    if (!SetCommTimeouts(hComm, &timeouts)) {
        printf("Could not set serial port timeouts\n");
        CloseHandle(hComm);
        return INVALID_HANDLE_VALUE;
    }
    return hComm;
}

int getDateTime() {
    // Get current date and time
    time_t raw_time;
    struct tm *tm_info;

    time(&raw_time);
    tm_info = localtime(&raw_time);

    // Store date and time in BMSData struct
    snprintf(bmsData.dateTime, sizeof(bmsData.dateTime), "%04d-%02d-%02d %02d:%02d:%02d", 
             tm_info->tm_year + 1900, 
             tm_info->tm_mon + 1, 
             tm_info->tm_mday, 
             tm_info->tm_hour, 
             tm_info->tm_min, 
             tm_info->tm_sec);

    printf("\n\nCurrent Time: %s\n", bmsData.dateTime);

}


int getBMSData(HANDLE hComm, int requestType) {
    const int REQ_LEN = 13;
    const unsigned char REQUEST_TOTAL_VOLTAGE_CURRENT_SOC[13] = {0XA5, 0X40, 0X90, 0X08, 0X00, 0X00, 0x00, 0x00, 0X00, 0X00, 0X00, 0X00, 0x7D};
    const unsigned char REQUEST_HIGHEST_LOWEST_VOLTAGE[13] = {0XA5, 0X40, 0X91, 0X08, 0X00, 0X00, 0x00, 0x00, 0X00, 0X00, 0X00, 0X00, 0x7E};
    const unsigned char REQUEST_MAX_MIN_TEMP[13] = {0XA5, 0X40, 0X92, 0X08, 0X00, 0X00, 0x00, 0x00, 0X00, 0X00, 0X00, 0X00, 0x7F};
    const unsigned char REQUEST_CHARGE_DISCHARGE_MOS_STATUS[13] = {0XA5, 0X40, 0X93, 0X08, 0X00, 0X00, 0x00, 0x00, 0X00, 0X00, 0X00, 0X00, 0x80};
    const unsigned char REQUEST_STATUS_INFO_1[13] = {0XA5, 0X40, 0X94, 0X08, 0X00, 0X00, 0x00, 0x00, 0X00, 0X00, 0X00, 0X00, 0x81};
    const unsigned char REQUEST_SINGLE_CELL_VOLTAGE[13] = {0XA5, 0X40, 0X95, 0X08, 0X00, 0X00, 0x00, 0x00, 0X00, 0X00, 0X00, 0X00, 0x82};
    const unsigned char REQUEST_SINGLE_CELL_TEMP[13] = {0XA5, 0X40, 0X96, 0X08, 0X00, 0X00, 0x00, 0x00, 0X00, 0X00, 0X00, 0X00, 0x83};
    const unsigned char REQUEST_SINGLE_CELL_BALANCE_STATUS[13] = {0XA5, 0X40, 0X97, 0X08, 0X00, 0X00, 0x00, 0x00, 0X00, 0X00, 0X00, 0X00, 0x84};
    const unsigned char REQUEST_SINGLE_CELL_FAILURE_STATUS[13] = {0XA5, 0X40, 0X98, 0X08, 0X00, 0X00, 0x00, 0x00, 0X00, 0X00, 0X00, 0X00, 0x85};

    const unsigned char *pRequest;
    unsigned char pResponse[300];

    switch( requestType ) { 
        case READ_BAT_TOTAL_VOLTAGE_CURRENT_SOC:
            pRequest = REQUEST_TOTAL_VOLTAGE_CURRENT_SOC;
            break;
        case READ_BAT_HIGHEST_LOWEST_VOLTAGE:
            pRequest = REQUEST_HIGHEST_LOWEST_VOLTAGE;
            break;
        case READ_BAT_MAX_MIN_TEMP:
            pRequest = REQUEST_MAX_MIN_TEMP;
            break;
        case READ_BAT_CHARGE_DISCHARGE_MOS_STATUS:
            pRequest = REQUEST_CHARGE_DISCHARGE_MOS_STATUS;
            break;
        case READ_BAT_STATUS_INFO_1:
            pRequest = REQUEST_STATUS_INFO_1;
            break;
        case READ_BAT_SINGLE_CELL_VOLTAGE:
            pRequest = REQUEST_SINGLE_CELL_VOLTAGE;
            break;
        case READ_BAT_SINGLE_CELL_TEMP:
            pRequest = REQUEST_SINGLE_CELL_TEMP;
            break;
        case READ_BAT_SINGLE_CELL_BALANCE_STATUS:
            pRequest = REQUEST_SINGLE_CELL_BALANCE_STATUS;
            break;
        case READ_BAT_SINGLE_CELL_FAILURE_STATUS:
            pRequest = REQUEST_SINGLE_CELL_FAILURE_STATUS;
            break;
        default:
            return 0;
    }

    const int MAX_RETRY = 1;
    DWORD bytesWritten;

    for (int i = 0; i < MAX_RETRY; i++) {
        
        BOOL status = WriteFile(hComm, pRequest, REQ_LEN, &bytesWritten, NULL);

        if (status) {
            printf("Data written to port, %lu bytes\n", bytesWritten);
        } else {
            printf("Could not write data to port\n");
        }


        DWORD bytesRead;
        status = ReadFile(hComm, pResponse, sizeof(pResponse), &bytesRead, NULL);

        if (status) {
            printf("Data read from port: ");
            for (int i = 0; i < bytesRead; i++) {
                printf("%02X ", pResponse[i]);
            }
            printf("\n");
        } else {
            printf("Could not read data from port\n");
        }

        switch( pResponse[2] ) {
            case READ_BAT_TOTAL_VOLTAGE_CURRENT_SOC:
                parseBmsResponseSoc(pResponse);
                break;
            case READ_BAT_HIGHEST_LOWEST_VOLTAGE:
                parseBmsResponseHighestLowestVoltage(pResponse);
                break;
            case READ_BAT_MAX_MIN_TEMP:
                parseBmsResponseMaxMinTemp(pResponse);
                break;
            case READ_BAT_CHARGE_DISCHARGE_MOS_STATUS:
                parseBmsResponseChargeDischargeMosStatus(pResponse);
                break;
            case READ_BAT_STATUS_INFO_1:
                parseBmsResponseStatusInfo1(pResponse);
                break;
            case READ_BAT_SINGLE_CELL_VOLTAGE:
                parseBmsResponseSingleCellVoltage(pResponse);
                break;
            case READ_BAT_SINGLE_CELL_TEMP:
                parseBmsResponseSingleCellTemp(pResponse);
                break;
            case READ_BAT_SINGLE_CELL_BALANCE_STATUS:
                parseBmsResponseSingleCellBalancingStatus(pResponse);
                break;
            case READ_BAT_SINGLE_CELL_FAILURE_STATUS:
                parseBmsResponseBatteryFailureStatus(pResponse);
                break;
            default:
                return 0;
            }
    }

}

int parseBmsResponseSoc(unsigned char *pResponse) {
    // Data bits start at index 4
    float cumulative_total_voltage = ((pResponse[4] << 8) + pResponse[5]) * 0.1f;
    float collect_total_voltage = ((pResponse[6] << 8) + pResponse[7]) * 0.1f;
    float current = (((pResponse[8] << 8) + pResponse[9]) - 30000) * 0.1f;
    float soc = ((pResponse[10] << 8) + pResponse[11]) * 0.1f;

    bmsData.voltage = cumulative_total_voltage;
    bmsData.current = current;
    bmsData.stateOfCharge = soc;

    printf("cumulative_total_voltage: %.2fV\n", cumulative_total_voltage);
    printf("collect_total_voltage: %.2fV\n", collect_total_voltage);
    printf("current: %.2fA\n", current);
    printf("soc: %.2f%%\n", soc);
}

int parseBmsResponseHighestLowestVoltage(unsigned char *pResponse) {
    // Data bits start at index 4
    float highest_single_voltage = (pResponse[4] << 8) + pResponse[5];
    int highest_voltage_cell_number = pResponse[6];
    float lowest_single_voltage = (pResponse[7] << 8) + pResponse[8];
    int lowest_voltage_cell_number = pResponse[9];

    bmsData.highestCellVoltage = highest_single_voltage;
    bmsData.lowestCellVoltage = lowest_single_voltage;

    printf("highest_single_voltage: %.2fmV\n", highest_single_voltage);
    printf("highest_voltage_cell_number: %d\n", highest_voltage_cell_number);
    printf("lowest_single_voltage: %.2fmV\n", lowest_single_voltage);
    printf("lowest_voltage_cell_number: %d\n", lowest_voltage_cell_number);
}

int parseBmsResponseMaxMinTemp(unsigned char *pResponse) {
    // Data bits start at index 4
    float max_temp = pResponse[4] - 40;
    int max_temp_cell_number = pResponse[5];
    float min_temp = pResponse[6] - 40;
    int min_temp_cell_number = pResponse[7];

    printf("max_temp: %.2fC\n", max_temp);
    printf("max_temp_cell_number: %d\n", max_temp_cell_number);
    printf("min_temp: %.2fC\n", min_temp);
    printf("min_temp_cell_number: %d\n", min_temp_cell_number);
}

int parseBmsResponseChargeDischargeMosStatus(unsigned char *pResponse) {
    // Data bits start at index 4
    int charge_discharge_status = pResponse[4];
    int mos_tube_charging_status = pResponse[5];
    int mos_tube_discharging_status = pResponse[6];
    int bms_life = pResponse[7];
    int remaining_capacity = (pResponse[8] << 24) | (pResponse[9] << 16) | (pResponse[10] << 8) | pResponse[11];

    bmsData.chargingDischargingStatus = charge_discharge_status;
    bmsData.chargingMOSStatus = mos_tube_charging_status;
    bmsData.dischargingMOSStatus = mos_tube_discharging_status;

    bmsData.remainingCapacity = remaining_capacity;

    printf("charge_discharge_status: %d\n", charge_discharge_status);
    printf("mos_tube_charging_status: %d\n", mos_tube_charging_status);
    printf("mos_tube_discharging_status: %d\n", mos_tube_discharging_status);
    printf("bms_life: %d\n", bms_life);
    printf("remaining_capacity: %dmAH\n", remaining_capacity);
}

int parseBmsResponseStatusInfo1(unsigned char *pResponse) {
    // Data bits start at index 4
    int battery_strings = pResponse[4];
    int number_of_temperature = pResponse[5];
    int charger_status = pResponse[6];
    int load_status = pResponse[7];
    int states = pResponse[8];

    int DI1_state = (states >> 0) & 1;
    int DI2_state = (states >> 1) & 1;
    int DI3_state = (states >> 2) & 1;
    int DI4_state = (states >> 3) & 1;
    int DO1_state = (states >> 4) & 1;
    int DO2_state = (states >> 5) & 1;
    int DO3_state = (states >> 6) & 1;
    int DO4_state = (states >> 7) & 1;

    g_number_of_battery_cells = battery_strings;
    g_number_of_temp_sensors = number_of_temperature;

    printf("battery_strings: %d\n", battery_strings);
    printf("number_of_temperature: %d\n", number_of_temperature);
    printf("charger_status: %d\n", charger_status);
    printf("load_status: %d\n", load_status);
    printf("DI1_state: %d\n", DI1_state);
    printf("DI2_state: %d\n", DI2_state);
    printf("DI3_state: %d\n", DI3_state);
    printf("DI4_state: %d\n", DI4_state);
    printf("DO1_state: %d\n", DO1_state);
    printf("DO2_state: %d\n", DO2_state);
    printf("DO3_state: %d\n", DO3_state);
    printf("DO4_state: %d\n", DO4_state);
}

int parseBmsResponseSingleCellVoltage(unsigned char *pResponse) {
    // Data bits start at index 4

    float cell_voltages[G_MAX_NUMBER_OF_CELLS];

    const int MESSAGE_LENGTH = 13;
    const int MAX_FRAMES = 16;
    const int CELLS_PER_FRAME = 3;

    int nCellsRead = 0; // number of battery cells read
    int readIndex = 4;

    for (int i = 0; i < MAX_FRAMES; i++) {
        if (nCellsRead == g_number_of_battery_cells) {
            break;
        }

        // check if frame number is correct
        int frame_number = pResponse[4 + i * MESSAGE_LENGTH];
        if (frame_number != i + 1) { // Frames are 1-indexed. i is 0-indexed
            printf("Frame number incorrect\n");
            continue;
        }

        readIndex = i * MESSAGE_LENGTH + 4 + 1; // 4 for the start flage, bms address, command, and data length. 1 for byte 0 being the frame serial number

        for (int j = 0; j < CELLS_PER_FRAME; j++) {
            if (nCellsRead == g_number_of_battery_cells) {
                break;
            }
            cell_voltages[nCellsRead] = (pResponse[readIndex] << 8) + pResponse[readIndex + 1];

            bmsData.cellVoltage[nCellsRead] = cell_voltages[nCellsRead];

            printf("cell_voltages[%d]: %.2fmV\n", nCellsRead, cell_voltages[nCellsRead]);
            nCellsRead++;
            readIndex += 2;
        }
    }
}

int parseBmsResponseSingleCellTemp(unsigned char *pResponse) {
    int cell_temps[G_MAX_NUMBER_OF_CELLS];

    const int MESSAGE_LENGTH = 13;
    const int MAX_FRAMES = 3;
    const int CELLS_PER_FRAME = 7;
    const int TEMPERATURE_OFFSET = 40;

    int nRead = 0; // number of temperature sensors read
    // Data bits start at index 4
    int readIndex = 4;

    for (int i = 0; i < MAX_FRAMES; i++) {
        if (nRead == g_number_of_temp_sensors) {
            break;
        }

        // check if frame number is correct
        int frame_number = pResponse[4 + i * MESSAGE_LENGTH];
        if (frame_number != i + 1) { // Frames are 1-indexed. i is 0-indexed
            printf("Frame number incorrect\n");
            continue;
        }

        readIndex = i * MESSAGE_LENGTH + 4 + 1; // 4 for the start flage, bms address, command, and data length. 1 for byte 0 being the frame serial number

        for (int j = 0; j < CELLS_PER_FRAME; j++) {
            if (nRead == g_number_of_temp_sensors) {
                break;
            }
            cell_temps[nRead] = pResponse[readIndex] - TEMPERATURE_OFFSET;

            bmsData.temperatures[nRead] = cell_temps[nRead];

            printf("cell_temps[%d]: %dC\n", nRead, cell_temps[nRead]);
            nRead++;
            readIndex++;
        }
    }
}

int parseBmsResponseSingleCellBalancingStatus(unsigned char *pResponse) {
    // Data bits start at index 4
    // Balance Status' are stored in reverse order. I.e. balance_status[0] is the balance status of the last cell

    boolean isBalancing = 0;

    for (int i = g_number_of_battery_cells - 1; i >= 0; i--) {
        isBalancing = isBalancing || pResponse[4 + i];
        bmsData.cellBalancingStatus[i] = pResponse[4 + i];
    }

    for (int i = 0; i < g_number_of_battery_cells; i++) {
        printf("bmsData.cellBalancingStatus[%d]: %d\n", i, bmsData.cellBalancingStatus[i]);
    }


    bmsData.balancingStatus = isBalancing;
}

int parseBmsResponseBatteryFailureStatus(unsigned char *pResponse) {
    // Data bits start at index 4
    // Alarms' are stored in reverse bit order. I.e. alarms[i][0] is the alarm of the ith byte, last bit

    char alarms[8][9];  // 8 bytes, each with 8 bits + null-terminator
    
    for (int i = 0; i < 8; ++i) {
        int byte = pResponse[4 + i];  // start from pResponse[4]
        char *currentString = alarms[i];
        
        for (int j = 7; j >= 0; --j) {  // loop through each bit in byte
            currentString[j] = ((byte >> j) & 1) ? '1' : '0';
        }
        currentString[8] = '\0';  // null-terminate the string
    }
    
    for (int i = 0; i < 8; ++i) {
        strcpy(bmsData.alarms[i], alarms[i]);
        printf("Binary string for byte %d: %s\n", i, alarms[i]);
    }
}

FILE *openCsvFile() {
    // Get current date and time
    time_t raw_time;
    struct tm *tm_info;

    time(&raw_time);
    tm_info = localtime(&raw_time);

    char dateTime[14];
    // Store date and time in dateTime array
    snprintf(dateTime, sizeof(dateTime), "%02d%02d%02d_%02d%02d%02d", 
             (tm_info->tm_year + 1900) % 100, 
             tm_info->tm_mon + 1, 
             tm_info->tm_mday, 
             tm_info->tm_hour, 
             tm_info->tm_min, 
             tm_info->tm_sec);


    char filePrefix[7] = "EPData";    
    char fileName[30];

    strcpy(fileName, filePrefix);  // Initialize fileName with filePrefix
    strcat(fileName, dateTime);    // Concatenate dateTime to fileName
    strcat(fileName, ".csv");      // Concatenate ".csv" to fileName

    FILE *fp = fopen(fileName, "w");

    if (fp == NULL) {
        printf("Could not open file for writing.\n");
        return NULL;
    }
    return fp;
}

int printCsvHeader(FILE *fp) {
    fprintf(fp, "Line #, Timestamp, Battery ID, Current (A), Voltage (V), State Of Charge (%%), Total Capacity, Remaining Capacity (mAH),");
    
    // Print cell voltage headers
    for (int i = 1; i <= G_MAX_NUMBER_OF_CELLS; ++i) {
        fprintf(fp, "Cell Voltage %d (mV),", i);
    }
    
    fprintf(fp, "Highest Cell Voltage (mV), Lowest Cell Voltage (mV),");
    
    // Print temperature headers
    for (int i = 1; i <= G_MAX_NUMBER_OF_TEMP_SENSORS; ++i) {
        fprintf(fp, "Temperature %d (C),", i);
    }
    
    fprintf(fp, "Charging (1) Discharging (2) Status, Charging MOS Status, Discharging MOS Status, Balancing Status, Cell Balancing Status,");

    // Print alarm headers
    for (int i = 1; i <= 8; ++i) {
        fprintf(fp, "Alarm %d,", i);
    }

    fprintf(fp, "\n");  // New line at the end
}

int outputBMSDataToCsv(FILE *fp) {
    // Write data
    fprintf(fp, "%d, %s, %d, %.2f, %.2f, %.2f, %.2f, %.2f, ", 
            bmsData.lineNumber,
            bmsData.dateTime, 
            bmsData.batteryID, 
            bmsData.current, 
            bmsData.voltage, 
            bmsData.stateOfCharge, 
            bmsData.totalCapacity, 
            bmsData.remainingCapacity);
    bmsData.lineNumber++;

    for (int i = 0; i < G_MAX_NUMBER_OF_CELLS; i++) {
        if (i >= g_number_of_battery_cells) {
            fprintf(fp, " , ");
        } else {
            fprintf(fp, "%.2f, ", bmsData.cellVoltage[i]);
        }
    }

    fprintf(fp, "%.2f, %.2f, ", 
            bmsData.highestCellVoltage, 
            bmsData.lowestCellVoltage);

    for (int i = 0; i < G_MAX_NUMBER_OF_TEMP_SENSORS; i++) {
        if (i >= g_number_of_temp_sensors) {
            fprintf(fp, " , ");
        } else {
            fprintf(fp, "%.2f, ", bmsData.temperatures[i]);
        }
    }

    fprintf(fp, "%d, %d, %d, %d, ", 
            bmsData.chargingDischargingStatus, 
            bmsData.chargingMOSStatus, 
            bmsData.dischargingMOSStatus, 
            bmsData.balancingStatus);

    // Combine cellBalancingStatus[i] into a single string
    char cellBalancingStr[G_MAX_NUMBER_OF_CELLS + 1]; // +1 for the null-terminator
    cellBalancingStr[G_MAX_NUMBER_OF_CELLS] = '\0'; // Null-terminate the string

    for (int i = 0; i < g_number_of_battery_cells; i++) {
        cellBalancingStr[i] = bmsData.cellBalancingStatus[i] + '0'; // '0' or '1'
    }

    // output the string to the csv file
    fprintf(fp, "\'%s\', ", cellBalancingStr);    

    for (int i = 0; i < 8; i++) {
        fprintf(fp, "\'%s\', ", bmsData.alarms[i]);
    }

    fprintf(fp, "\n");

    printf("Data written to csv file\n");

    return 0;
}


int main(int argc, char *argv[]) {
    bmsData.lineNumber = 1;
    int comPort = NO_COM_PORT_NUMBER_SUPPLIED;
    comPort = readProgramParams(argc, argv);
    if (comPort == INVALID_COM_PORT_NUMBER || comPort == INVALID_DELAY_TIME_SUPPLIED) {
        return 1; // Invalid COM port number or delay time supplied.
    }

    HANDLE hComm = setupCOMPort(comPort);

    if (hComm == INVALID_HANDLE_VALUE) return 1; // Could not open COM port.

    printf("Opening serial port successful\n");

    FILE *fp = openCsvFile();
    printCsvHeader(fp);

    while (1) {
        getDateTime();
        getBMSData(hComm, READ_BAT_TOTAL_VOLTAGE_CURRENT_SOC);
        getBMSData(hComm, READ_BAT_HIGHEST_LOWEST_VOLTAGE);
        getBMSData(hComm, READ_BAT_MAX_MIN_TEMP);
        getBMSData(hComm, READ_BAT_CHARGE_DISCHARGE_MOS_STATUS);
        getBMSData(hComm, READ_BAT_STATUS_INFO_1);
        getBMSData(hComm, READ_BAT_SINGLE_CELL_VOLTAGE);
        getBMSData(hComm, READ_BAT_SINGLE_CELL_TEMP);
        getBMSData(hComm, READ_BAT_SINGLE_CELL_BALANCE_STATUS);
        getBMSData(hComm, READ_BAT_SINGLE_CELL_FAILURE_STATUS);

        outputBMSDataToCsv(fp);
        Sleep(g_delay_time_ms);
    }

    // Close the file
    fclose(fp);
    // Close the COM port
    CloseHandle(hComm);
    return 0;
}
