import extractTransformLoad
import kpiExport

def main():
    print("Generating reports...")
    extractTransformLoad.main()
    kpiExport.main()
    print("Generation complete.")

if __name__ == "__main__":
    main()