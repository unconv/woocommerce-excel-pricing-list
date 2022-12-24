# WooCommerce Excel Pricing List Plugin

This is a WordPress plugin I made for a YouTube video about Chat GPT. I coded this plugin by giving prompts to Chat GPT.

The plugin adds a REST API endpoint `/wp-json/wc-xls-pricing-list/v1/products/xls` that returns an XLS file with all the product SKU's, names and prices from the WooCommerce store.

The pricing list is updated automatically when new products are added or modified, but a new version is only generated every 30 minutes.

Currently the plugin supports only regular products, not variable products. Only products that have a price will be added to the Excel sheet.

## Watch the video

You can watch the YouTube video I made here: https://www.youtube.com/watch?v=w3_KVbWYews
