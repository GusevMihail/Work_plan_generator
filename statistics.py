import matplotlib.pyplot as plt
import numpy as np
import seaborn as sns


def visualize_statistics(counter, title='counter'):
    """
    Visualizes counter

    counter: np array of ints, x axis - predicted class, y axis - actual class
                      [i][j] should have the count of samples that were predicted to be class i,
                      but have j in the ground truth

    """

    def append_unique(element_to_find, list_of_elements: list):
        if element_to_find not in list_of_elements:
            list_of_elements.append(element_to_find)

    xlabels = []
    ylabels = []
    for (xlabel, ylabel), count in counter.items():
        append_unique(xlabel, xlabels)
        append_unique(ylabel, ylabels)
    size = (len(ylabels), len(xlabels))

    counts = np.zeros(size, dtype=np.int)
    for (xlabel, ylabel), count in counter.items():
        x = xlabels.index(xlabel)
        y = ylabels.index(ylabel)
        counts[y, x] = count
    # Adapted from
    # https://stackoverflow.com/questions/2897826/confusion-matrix-with-number-of-classified-misclassified-instances-on-it-python
    fig, ax = plt.subplots()
    # fig = plt.figure(figsize=(10, 10))
    # plt.title("Confusion counter")
    # plt.ylabel("predicted")
    # plt.xlabel("ground truth")

    res = plt.imshow(counts, cmap='GnBu', interpolation='nearest')
    cb = fig.colorbar(res)
    plt.xticks(np.arange(size[1]), labels=xlabels)
    plt.yticks(np.arange(size[0]), labels=ylabels)
    plt.setp(ax.get_xticklabels(), rotation=45, ha="right", rotation_mode="anchor")
    plt.ylim(size[0] - .5, - .5)
    # plt.ylim(-.5, size[0] - .5) # сортировка снизу вверх
    for i, row in enumerate(counts):
        for j, count in enumerate(row):
            plt.text(j, i, int(count), fontsize=14, horizontalalignment='center', verticalalignment='center')

    # fig, ax = plt.subplots()
    # im = ax.imshow(counts)
    # ax.set_xticks(np.arange(size[1]))
    # ax.set_yticks(np.arange(size[0]))
    # ax.set_xticklabels(xlabels)
    # ax.set_yticklabels(ylabels)
    # fig.tight_layout(rect=(0,0,1,1))

    # counts = np.random.rand(10, 12)
    # ax = sns.heatmap(counts, linewidths=.2)
    plt.show()
