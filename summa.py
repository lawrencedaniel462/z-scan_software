import datetime

x = datetime.datetime.now()

print(x.year)
print(f"Z-Scan Sample analysis had done at SSN Research Center on {x.strftime("%d")}"
                        f" {x.strftime("%b")} {x.strftime("%Y")} {x.strftime("%X")}")