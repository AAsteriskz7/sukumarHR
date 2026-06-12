n = int(input())
line = input().split()
nums = []
for x in line:
    nums.append(int(x))
print(min(nums, key=lambda v: (abs(v), -v)))
