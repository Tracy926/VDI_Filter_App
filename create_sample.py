import pandas as pd
from datetime import datetime, timedelta
import random

# Create sample data
data = {
    'Associated User': [
        'john.smith', 'jane.doe', 'john.smith', 'mike.johnson', 'jane.doe',
        'sarah.wilson', 'john.smith', '', 'mike.johnson', 'sarah.wilson',
        'robert.brown', '', 'jane.doe', 'robert.brown', 'john.smith',
        'sarah.wilson', 'mike.johnson', 'peter.parker', 'jane.doe', 'robert.brown'
    ],
    'Session Start Time': [
        '12/18/2025 2:29:58 PM', '12/17/2025 9:15:30 AM', '12/15/2025 1:45:22 PM',
        '12/14/2025 8:30:45 AM', '12/13/2025 3:22:10 PM', '12/12/2025 10:05:33 AM',
        '12/11/2025 2:15:50 PM', '12/10/2025 11:45:20 AM', '12/09/2025 4:32:15 PM',
        '12/08/2025 9:10:25 AM', '12/07/2025 1:55:40 PM', '12/06/2025 8:20:10 AM',
        '12/05/2025 3:40:50 PM', '12/04/2025 10:15:30 AM', '12/03/2025 2:25:15 PM',
        '12/02/2025 11:50:45 AM', '12/01/2025 5:15:20 PM', '11/28/2025 9:30:10 AM',
        '11/25/2025 1:45:55 PM', '11/20/2025 8:10:40 AM'
    ],
    'Session End Time': [
        '12/18/2025 3:45:30 PM', '12/17/2025 10:30:15 AM', '12/15/2025 2:50:45 PM',
        '12/14/2025 9:15:20 AM', '12/13/2025 4:35:50 PM', '12/12/2025 11:20:10 AM',
        '12/11/2025 3:20:25 PM', '12/10/2025 12:30:40 PM', '12/09/2025 5:45:30 PM',
        '12/08/2025 10:25:15 AM', '12/07/2025 2:40:20 PM', '12/06/2025 9:10:50 AM',
        '12/05/2025 4:55:25 PM', '12/04/2025 11:30:45 AM', '12/03/2025 3:40:55 PM',
        '12/02/2025 12:35:20 PM', '12/01/2025 6:20:40 PM', '11/28/2025 10:45:30 AM',
        '11/25/2025 2:55:30 PM', '11/20/2025 9:25:15 AM'
    ],
    'Endpoint Name': [
        'ENDPOINT-001', 'ENDPOINT-002', 'ENDPOINT-001', 'ENDPOINT-003', 'ENDPOINT-002',
        'ENDPOINT-004', 'ENDPOINT-001', 'ENDPOINT-005', 'ENDPOINT-003', 'ENDPOINT-004',
        'ENDPOINT-002', '', 'ENDPOINT-001', 'ENDPOINT-003', 'ENDPOINT-002',
        'ENDPOINT-004', 'ENDPOINT-005', 'ENDPOINT-001', 'ENDPOINT-003', 'ENDPOINT-002'
    ],
    'Machine Name': [
        'VDI-MACHINE-01', 'VDI-MACHINE-02', 'VDI-MACHINE-01', 'VDI-MACHINE-03', 'VDI-MACHINE-04',
        'VDI-MACHINE-05', 'VDI-MACHINE-01', 'VDI-MACHINE-06', 'VDI-MACHINE-03', 'VDI-MACHINE-05',
        'VDI-MACHINE-02', 'VDI-MACHINE-07', 'VDI-MACHINE-01', 'VDI-MACHINE-03', 'VDI-MACHINE-04',
        'VDI-MACHINE-05', 'VDI-MACHINE-06', 'VDI-MACHINE-02', 'VDI-MACHINE-04', 'VDI-MACHINE-07'
    ],
    'Delivery Group Name': [
        'Marketing', 'Engineering', 'Marketing', 'Sales', 'Finance',
        'HR', 'Marketing', 'Operations', 'Sales', 'HR',
        'Engineering', 'Finance', 'Marketing', 'Sales', 'Engineering',
        'HR', 'Operations', 'Marketing', 'Finance', 'Sales'
    ],
    'IP Address': [
        '192.168.1.101', '192.168.1.102', '192.168.1.103', '192.168.1.104', '192.168.1.105',
        '192.168.1.106', '', '192.168.1.108', '192.168.1.109', '192.168.1.110',
        '192.168.1.111', '192.168.1.112', '', '192.168.1.114', '192.168.1.115',
        '192.168.1.116', '192.168.1.117', '192.168.1.118', '192.168.1.119', ''
    ],
    'Associated User Display Name': [
        'John Smith', 'Jane Doe', 'John Smith', 'Mike Johnson', 'Jane Doe',
        'Sarah Wilson', 'John Smith', '', 'Mike Johnson', 'Sarah Wilson',
        'Robert Brown', 'Robert Brown', '', 'Robert Brown', 'John Smith',
        'Sarah Wilson', 'Mike Johnson', 'Peter Parker', 'Jane Doe', ''
    ]
}

# Create DataFrame
df = pd.DataFrame(data)

# Save to Excel
output_file = 'sample_input.xlsx'
df.to_excel(output_file, index=False, sheet_name='Session Data')

print(f"✓ Sample input file created: {output_file}")
print(f"  Total rows: {len(df)}")
print(f"  Columns: {list(df.columns)}")
print("\nYou can now use this file to test the VDI User Filter App!")
print("Run: python run.py")
