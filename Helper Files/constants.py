AMAZON_KEYS = {
    'order-id' : 'order-item-id',
    'secondary-order-id' : 'order-id',
    'same-buyer-order-id' : 'order-id',
    'purchase-date' : 'purchase-date',
    'payments-date' : 'payments-date',
    'buyer-email' : 'buyer-email',
    'buyer-name' : 'buyer-name',
    'buyer-phone-number' : 'buyer-phone-number',
    'sku' : 'sku',
    'title' : 'product-name',
    'quantity-purchased' : 'quantity-purchased',
    'currency' : 'currency',
    'item-price' : 'item-price',
    'item-tax' : 'item-tax',
    'shipping-price' : 'shipping-price',
    'shipping-tax' : 'shipping-tax',
    'ship-service-level' : 'ship-service-level',
    'recipient-name' : 'recipient-name',
    'ship-address-1' : 'ship-address-1',
    'ship-address-2' : 'ship-address-2',
    'ship-address-3' : 'ship-address-3',
    'ship-city' : 'ship-city',
    'ship-state' : 'ship-state',
    'ship-postal-code' : 'ship-postal-code',
    'ship-country' : 'ship-country',
    'ship-phone-number' : 'ship-phone-number',
    'delivery-start-date' : 'delivery-start-date',
    'delivery-end-date' : 'delivery-end-date',
    'delivery-time-zone' : 'delivery-time-zone',
    'delivery-Instructions' : 'delivery-Instructions',
    'sales-channel' : 'sales-channel',
}

AMAZON_WAREHOUSE_KEYS = {
    'order-id' : 'Shipment Item ID',
    'secondary-order-id' : 'Amazon Order Id',
    'purchase-date' : 'Purchase Date',
    'payments-date' : 'Payments Date',
    'buyer-email' : 'Buyer E-mail',
    'buyer-name' : 'Buyer Name',
    'buyer-phone-number' : 'Buyer Phone Number',
    'sku' : 'Merchant SKU',
    'title' : 'Title',
    'quantity-purchased' : 'Dispatched Quantity',
    'currency' : 'Currency',
    'item-price' : 'Item Price',
    'item-tax' : 'Item Tax',
    'shipping-price' : 'Delivery Price',
    'shipping-tax' : 'Delivery Tax',
    'ship-service-level' : 'Delivery Service Level',
    'recipient-name' : 'Recipient Name',
    'ship-address-1' : 'Delivery Address 1',
    'ship-address-2' : 'Delivery Address 2',
    'ship-address-3' : 'Delivery Address 3',
    'ship-city' : 'Delivery City/Town',
    'ship-state' : 'Delivery County',
    'ship-postal-code' : 'Delivery Postcode',
    'ship-country' : 'Delivery Country Code',
    'ship-phone-number' : 'Delivery Phone Number',
    'sales-channel' : 'Sales Channel',
}

SALES_CHANNEL_PROXY_KEYS = {'AmazonCOM': AMAZON_KEYS, 'AmazonEU': AMAZON_KEYS, 'Amazon Warehouse': AMAZON_WAREHOUSE_KEYS,}

# Value corresponds to proxy_keys
TEMPLATE_SHEET_MAPPING= {
        'Order ID' : 'secondary-order-id',
        'Unique ID' : 'order-id',
        'Repicient' : 'recipient-name',
        'Cur.' : 'currency',
        'Q-ty' : 'quantity-purchased',
        'Price' : 'item-price',
        'Tax' : 'item-tax',
        'Shipping' : 'shipping-price',
        'Shipping Tax' : 'shipping-tax',
        'Purchase' : 'purchase-date',
        'Payment' : 'payments-date',
        'Country' : 'ship-country'
}

EU_SUMMARY_HEADERS = ['Currency',
                '  Date ',
                '  Total',
                'Total #',
                'NON-VAT',
                '  #',
                '',
                'UK TAXES',
                'NON-VAT -Taxes',
                '',]


COM_SUMMARY_HEADERS = ['Currency',
                '  Date ',
                '  Total',
                'Total #',
                'EU   ',
                '  #',
                'NON-EU',
                '  #',
                '',
                'TAXES    ']

VBA_ERROR_ALERT = 'ERROR_CALL_DADDY'
VBA_KEYERROR_ALERT = 'ERROR_IN_SOURCE_HEADERS'
VBA_OK = 'EXPORTED_SUCCESSFULLY'
VBA_NO_NEW_JOB = 'NO NEW JOB'