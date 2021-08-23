import json
import boto3
import openpyxl
from openpyxl import Workbook
from botocore.config import Config

class aws_inventory:
    def __init__(self, ec2_config):
        self.ec2_config = ec2_config
    
    def get_instances(self):
        instanceData = self.ec2_config.describe_instances()
        #print(instanceData)
        return instanceData

    def get_vpcs(self):
        vpcData = self.ec2_config.describe_vpcs()
        # print(vpcData)
        return vpcData

    def get_subnets(self):
        subnetData = self.ec2_config.describe_subnets()
        # print(subnetData)
        return subnetData

    def build_dict(self):
        dataDict = {
            'Instance Data': self.get_instances(),
            'VPC Data': self.get_vpcs(),
            'Subnet Data': self.get_subnets()
            }
        return dataDict
    
    def get_instance_tags(self, data):
        required_tags = ["FISMA ID", "Name", "Operating System"]
        tagList = []
        for item in required_tags:
            tagList.append("N/A")
        for tag in data["Instances"][0]["Tags"]:
            for index, item in enumerate(required_tags):
                if tag["Key"] == item:
                    tagList[index] = tag["Value"]
        return tagList

    def get_instance_subnet(self, data, subnetDict):
        required_info = ["Subnet ID", "Subnet", "Subnet CIDR"]
        subnetInfo = []
        for item in required_info:
            subnetInfo.append("N/A")
        subnetInfo[0] = data["Instances"][0]["SubnetId"]
        for item in subnetDict["Subnets"]:
            if item["SubnetId"] == subnetInfo[0]:
                subnetInfo[1] = item["Tags"][0]["Value"]
                subnetInfo[2] = item["CidrBlock"]
        return subnetInfo

    def write_ec2_data(self, dataDict, workbook):
        ws = workbook.active
        ws.title = "EC2 Data"
        # data = json.loads(json.dumps(dataDict['Instance Data'], indent=4, default=str))
        headerList = ["Instance Name", "Instance ID", "FISMA ID", "IP Address", "AMI ID", "Operating System", "VPC", "VPC ID", "Subnet", "Subnet CIDR", "Subnet ID"]
        tableData = []
        ws.append(headerList)
        data = dataDict["Instance Data"]
        i = 0
        while i != len(data["Reservations"]):
            if data["Reservations"][i]["Instances"][0]["State"]["Name"] != "terminated":
                vpc = "N/A"
                instanceId = data["Reservations"][i]["Instances"][0]["InstanceId"]
                amiId = data["Reservations"][i]["Instances"][0]["ImageId"]
                privateIpAddress = data["Reservations"][i]["Instances"][0]["PrivateIpAddress"]
                vpcId = data["Reservations"][i]["Instances"][0]["VpcId"]
                
                tagList = self.get_instance_tags(data["Reservations"][i])
                subnetInfo = self.get_instance_subnet(data["Reservations"][i], dataDict["Subnet Data"])
                # for item in dataDict["Subnet Data"]["Subnets"]:
                #     if item["SubnetId"] == subnetId:
                #         subnet = item["Tags"][0]["Value"]
                #         subnetBlock = item["CidrBlock"]
                        

                ws.append([tagList[1], instanceId, tagList[0], privateIpAddress, amiId, tagList[2], vpc, vpcId, subnetInfo[1], subnetInfo[2], subnetInfo[0]])
            i+=1

        # data = json.loads(json.dumps(v, indent=4, default=str))

    def compile(self):
        workbook = Workbook()
        dataDict = self.build_dict()
        self.write_ec2_data(dataDict, workbook)
        workbook.save(filename="aws_inventory.xlsx")
        
    def output_json(self):
        dataDict = self.build_dict()
        for k,v in dataDict.items():
            with open(k+'.json','w+') as f:
                json.dump(v, f, indent = 4, default=str)
                f.close()
        


def main():
    for region in ['us-east-1']:
        sub_config = Config(
            region_name = region
        )
        ec2_config = boto3.client('ec2', config=sub_config)
        inventory = aws_inventory(ec2_config)
        inventory.output_json()
        inventory.compile()

if __name__ == "__main__":
  main()
