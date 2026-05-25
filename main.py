import instr
from xls_writer import write_to_file

numb_meas = 10
results = []

visa_devices = instr.scan_visa_devices()
lan_devices = instr.scan_lan_devices("172.16.4", 20, 30)

devices = lan_devices + visa_devices 

print("Dostępne urządzenia:")
for i, dev in enumerate(devices):
    print(i, dev["addr"], dev["idn"])

if not devices:
    print("Nie wykryto urządzenia.")
    exit()

instrument = instr.connect(devices[1]["addr"])
print(instrument.query("*IDN?").strip())
try:
    instr.write_command(instrument, "CONF:VOLT:DC 1")

    for x in range(numb_meas):
        value = instr.read_value(instrument, "READ?")
        results.append(value)
        print(x + 1, f"{value:.10f}")

    write_to_file(results, "pomiary.xlsx")

finally:
    instr.disconnect(instrument)