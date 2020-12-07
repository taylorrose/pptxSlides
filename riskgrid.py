import matplotlib.pyplot as plt
import matplotlib.image as mpimg
import matplotlib


def risk_score(severity, likelihood):
    r = str(severity) + str(likelihood)
    if r == "33":
        return 9
    if r == "32":
        return 8
    if r == "23":
        return 7
    if r == "31":
        return 6
    if r == "22":
        return 5
    if r == "13":
        return 4
    if r == "21":
        return 3
    if r == "12":
        return 2
    if r == "11":
        return 1

#TODO: Cache Images in Image Folder
def plot_risk(severity, liklihood, Mseverity, Mliklihood, name):
    matplotlib.rcParams['font.sans-serif'] = "Avenir"

    fig, ax = plt.subplots(figsize=(6, 6))
    img = mpimg.imread('img/grid.png')

    plt.imshow(img)
    plt.xlabel('Likelihood', fontsize=22)
    plt.ylabel('Severity', fontsize=22)
    ax.yaxis.label.set_color('black')
    ax.xaxis.label.set_color('black')
    plt.tick_params(
        which='both',
        bottom=False,
        left=False,
        top=False,
        labelbottom=False,
        labelleft=False)
    ax.spines['bottom'].set_color('white')
    ax.spines['top'].set_color('white')
    ax.spines['right'].set_color('white')
    ax.spines['left'].set_color('white')

    x = [0, 100, 200, 300, 400, 500, 600, 700, 800, 900]
    y = [0, 100, 200, 300, 400, 500, 600, 700, 800, 900]

    plt.scatter(x, y, s=1, c='blue', marker='o', alpha=0)

    x = Mliklihood
    y = Mseverity

    plt.scatter(x * 300 - 150, abs((y * 300 - 150) - 900), s=5000, c='#4374FF', marker='o', alpha=0.3)
    plt.scatter(x * 300 - 150, abs((y * 300 - 150) - 900), s=1400, c='#4374FF', marker='o', alpha=1)

    x = liklihood
    y = severity

    plt.scatter(x * 300 - 150, abs((y * 300 - 150) - 900), s=5000, c='#FD014E', marker='o', alpha=0.3)
    plt.scatter(x * 300 - 150, abs((y * 300 - 150) - 900), s=1400, c='#FD014E', marker='o', alpha=1)


    plt.savefig('img/risks/' + name + '.png', dpi=200)
    plt.show()
