# The Helper Tools for Deploy ACI From Excel (DAFE)
## Description

The DAFE software is a tool used by Cisco CX Team for quick ACI deployment, but it is not open for customer use per se, and also its Excel datasheet requires line-by-line editing and is not utilized by large-scale users for high-volume deployment. I have prepared some tools for customers who have these requirements.

The data_process tool is used to generate data for DAFE's Excel sheet in bulk, which is suitable for customers with large deployments, such as 200 EPGs during one deployment.

The data_deploy tool is used to complete the actual ACI deployment with the Excel data. I use the Cobra SDK to make it easier for network admin to understand. And I open it to the customer, so that the customer can process it again for their own network operation and maintenance.

Currently I have developed the framework and several ACI configure functions. Subsequence config components will be added as the python functions in the future.