import re

text = """
Tracked 48

No Signature

Delivered By

Postage on Account GB

Parcel

2kg

Q23

A2

32-071 636 0000-5F7 D0A A74

OF 2242 7598 5GB

Return Address

Cavear Ltd

228 Flemington

Street

Glasgow

G21 4BF

Customer Ref: 

10207 / 204-4846383-6917168

2x 11" Silver Shade

Royal Mail: UK's lowest average parcel carbon footprint 200g CO2e

Nina rokitka

FLAT 18 ALBERT MCKENZIE 

HOUSE

100 SPA ROAD

LONDON

SE16 3QT
"""

# Use regex to find the customer reference
match = re.search(r'Customer Ref:\s*(\d+)\s*/\s*([0-9A-Za-z-]+)', text)

if match:
    customer_ref = match.group(2)  # Get the second capturing group
    print(customer_ref)  # Output: 204-4846383-6917168
