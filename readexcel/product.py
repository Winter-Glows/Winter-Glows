from base import *


class Product4Sample(Base):

    def __init__(self):
        super().__init__()
        self.msg = ''

    def get_baseinfo(self, zkid):
        """获取样本基本信息，家族信息，产品信息"""
        result = {}
        tmp = self.do_request(url_key='order', method='get', params={"sample__barcode__icontains": zkid})['results'][0]
        pprint(tmp)
        is_pedigree = tmp['order_orderproduct_info'][0]['product_info']['is_pedigree']
        if is_pedigree:
            result.update({"pedigree_id": tmp['pedigree_info'][0]['number']})
            result.update({"is_pedigree": is_pedigree})
            id = tmp['pedigree_info'][0]['id']
            tmp2 = self.do_request(url_key='pedigreeperson', method='get', params={"pedigree__id": id})['results']
            temp_list = [ "family_" + str(i) for i in range(5,0,-1)]
            for record in tmp2:
                if record['sample_info']['barcode'] == zkid:
                    # print("=====")
                    pprint(record)
                    result.update({"zk_name": record['sample_info']['name']})
                    result.update({"birthday": record['sample_info']['birthday']})
                    result.update({"gender": record['sample_info']['gender_info']['name']})
                    result.update({"family_history": record['sample_info']['family_history']})
                    result.update({"medicine_history": record['sample_info']['medicine_history']})
                    result.update({"clinical_diagnosis": record['sample_info']['clinical_diagnosis']})
                    result.update({"family_history": record['sample_info']['family_history']})
                    result.update({"product": record['sample_info']['order_info'][0]['product_info'][0]['product_info']['name']})
                    result.update({"product_numner": record['sample_info']['order_info'][0]['product_info'][0]['product_info']['number']})
                    result.update({"hospital ": record['sample_info']['order_info'][0]['send_hospital']})
                    result.update({"doc ": record['sample_info']['order_info'][0]['send_doctor']})
                    result.update({"hos_num ": record['sample_info']['order_info'][0]['hospital_number']})
                    result.update({"sample_type ": record['sample_info']['internal_samples_info'][0]['sample_type_info']['name']})
                    result.update({"sampling_time ": record['sample_info']['internal_samples_info'][0]['sampling_time']})
                    result.update({"receive_time ": record['sample_info']['internal_samples_info'][0]['receive_time']})
                    result.update({"nation ": record['sample_info']['nation_info']['name']})
                else:
                    temp_key = temp_list.pop()
                    #pprint(record)
                    result.update({temp_key: record["person_type_info"]["name"]})
                    result.update({temp_key + "_name": record['sample_info']['name']})
                    result.update({temp_key + "_gender": record['sample_info']['gender_info']['name']})
                    result.update({temp_key + "_id": record['sample_info']['barcode']})
        else:
            tmp3 = self.do_request(url_key='sample', method='get', params={"barcode__icontains": zkid})['results'][0]
            # pprint(tmp3)
            result.update({"zk_name": tmp3['internal_samples_info'][0]['sample_info']['name']})
            result.update({"age": tmp3['age']})
            result.update({"birthday": tmp3['birthday']})
            result.update({"product": []})
            for i in tmp3['internal_samples_info'][0]['product_info']:
                result["product"].append(i['name'])
            result.update({"sample_type ": tmp3['internal_samples_info'][0]['sample_type_info']['name']})
            result.update({"sampling_time ": tmp3['internal_samples_info'][0]['sampling_time']})
            result.update({"receive_time ": tmp3['internal_samples_info'][0]['receive_time']})
            result.update({"nation ": tmp3['nation_info']['name']})
        return result

if __name__ == "__main__":
    #demo
    from pprint import pprint
    test = Product4Sample()
    test_id = "320301666100"
    print(f"test: id={test_id}")
    res = test.get_baseinfo(test_id)
    pprint(res)
    print(len(res))