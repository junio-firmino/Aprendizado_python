from tqdm import tqdm
from time import sleep

for cont in tqdm(range(30), desc='teste', colour='red'):
    sleep(0.5)
