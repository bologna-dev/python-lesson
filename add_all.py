def add_all(nums):
    ret_num = 0
    for n in nums:
        ret_num += n 
    return(ret_num)

# main
all_nums = [1,2,3,4,5,6,7,8,9,10]
sum_num = add_all(all_nums)
print("合計は、" + str(sum_num) + "です")
