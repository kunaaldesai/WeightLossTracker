import pandas as pd
import matplotlib.pyplot as plt

def plot_weight_trend(filename):
    df = pd.read_excel(filename, engine='openpyxl')
    
    # Convert 'Date & Time' to datetime format
    df['Date & Time'] = pd.to_datetime(df['Date & Time'])
    
    # Plot Body Weight (lbs) trend
    plt.figure(figsize=(10, 5))
    plt.plot(df['Date & Time'], df['Body Weight (lbs)'], marker='o', label='Body Weight (lbs)')
    
    # Plot Target Weight if exists
    if 'Target Weight (lbs)' in df.columns:
        plt.axhline(y=df['Target Weight (lbs)'].iloc[0], color='r', linestyle='--', label='Target Weight')
    
    plt.title('Body Weight Trend')
    plt.xlabel('Date')
    plt.ylabel('Weight (lbs)')
    plt.legend()
    plt.grid(True)
    plt.tight_layout()
    plt.savefig("weight_trend.png")
    plt.show()

def plot_body_fat_trend(filename):
    df = pd.read_excel(filename, engine='openpyxl')

    # Convert 'Date & Time' to datetime format
    df['Date & Time'] = pd.to_datetime(df['Date & Time'])

    # Plot Body Fat % trend
    plt.figure(figsize=(10, 5))
    plt.plot(df['Date & Time'], df['Body Fat %'], marker='o', color='g', label='Body Fat %')
    
    plt.title('Body Fat Percentage Trend')
    plt.xlabel('Date')
    plt.ylabel('Body Fat %')
    plt.legend()
    plt.grid(True)
    plt.tight_layout()
    plt.savefig("body_fat_trend.png")
    plt.show()

def main():
    filename = "weight_loss_tracker.xlsx"

    while True:
        action = input("Bodyweight / Bodyfat / Exit ? ").lower()
        if action == "bodyweight":
            plot_weight_trend(filename)
        elif action == "bodyfat":
            plot_body_fat_trend(filename)
        elif action == "exit":
            break
    
    
    

if __name__ == "__main__":
    main()
