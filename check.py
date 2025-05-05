Based on the diagnostic output showing TLS version incompatibility and firewall concerns, I'd like to help you troubleshoot these specific issues. Let's create a more focused script to test and fix the TLS version issue first, as this is often the root cause when proxy settings don't solve the problem.

```python
import requests
import ssl
import socket
import sys
import urllib3
import certifi

def check_tls_compatibility():
    print("=" * 60)
    print(" CHECKING TLS COMPATIBILITY ".center(60, "="))
    print("=" * 60)
    
    # Print Python and OpenSSL version information
    print(f"Python version: {sys.version}")
    print(f"OpenSSL version: {ssl.OPENSSL_VERSION}")
    print(f"Default TLS version: {ssl.get_default_verify_paths()}")
    print(f"Certifi version: {certifi.__version__}")
    print(f"Requests version: {requests.__version__}")
    print(f"Urllib3 version: {urllib3.__version__}")
    
    # Test different TLS versions
    print("\nTesting different TLS configurations:")
    
    urls = ["https://www.google.com", "https://www.howsmyssl.com/a/check"]
    
    # Default TLS configuration
    print("\n1. Testing with default TLS configuration...")
    try:
        response = requests.get(urls[0], timeout=5, proxies={'http': None, 'https': None})
        print(f"✅ Success with default TLS! Status code: {response.status_code}")
    except Exception as e:
        print(f"❌ Failed with default TLS: {e}")
    
    # Force TLS 1.2
    print("\n2. Testing with forced TLS 1.2...")
    try:
        session = requests.Session()
        session.mount('https://', urllib3.PoolManager(
            cert_reqs='CERT_REQUIRED',
            ca_certs=certifi.where(),
            ssl_version=ssl.PROTOCOL_TLSv1_2
        ))
        response = session.get(urls[0], timeout=5, proxies={'http': None, 'https': None})
        print(f"✅ Success with TLS 1.2! Status code: {response.status_code}")
    except Exception as e:
        print(f"❌ Failed with TLS 1.2: {e}")
    
    # Check detailed TLS info
    print("\n3. Checking detailed TLS information...")
    try:
        response = requests.get(urls[1], timeout=5, proxies={'http': None, 'https': None})
        if response.status_code == 200:
            tls_info = response.json()
            print(f"✅ TLS check successful!")
            print(f"TLS Version: {tls_info.get('tls_version', 'Unknown')}")
            print(f"Cipher Suite: {tls_info.get('cipher_suite', 'Unknown')}")
            print(f"Rating: {tls_info.get('rating', 'Unknown')}")
            return tls_info.get('tls_version', 'Unknown')
    except Exception as e:
        print(f"❌ Failed to get TLS details: {e}")
    
    return None

def check_firewall_interference():
    print("\n" + "=" * 60)
    print(" CHECKING FIREWALL INTERFERENCE ".center(60, "="))
    print("=" * 60)
    
    # Test if specific ports commonly used by Python are blocked
    ports_to_test = [80, 443, 8080]
    results = {}
    
    for port in ports_to_test:
        try:
            if port == 80:
                host = "http://httpbin.org/get"
            else:
                host = "https://httpbin.org/get"
            
            print(f"\nTesting connection to port {port}...")
            response = requests.get(
                host, 
                proxies={'http': None, 'https': None},
                timeout=5
            )
            print(f"✅ Port {port} is accessible. Status code: {response.status_code}")
            results[port] = True
        except Exception as e:
            print(f"❌ Port {port} might be blocked: {e}")
            results[port] = False
    
    # Check if Python process is being blocked specifically
    print("\nTesting if Python is being specifically targeted...")
    
    user_agents = [
        # Default Python user agent
        "python-requests/2.28.1",
        # Browser-like user agent
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    ]
    
    ua_results = {}
    for ua in user_agents:
        try:
            print(f"\nTesting with User-Agent: {ua}")
            headers = {'User-Agent': ua}
            response = requests.get(
                "https://www.google.com", 
                headers=headers,
                proxies={'http': None, 'https': None},
                timeout=5
            )
            print(f"✅ Success with {ua}! Status code: {response.status_code}")
            ua_results[ua] = True
        except Exception as e:
            print(f"❌ Failed with {ua}: {e}")
            ua_results[ua] = False
    
    return results, ua_results

def provide_solutions(tls_version, port_results, ua_results):
    print("\n" + "=" * 60)
    print(" SOLUTIONS ".center(60, "="))
    print("=" * 60)
    
    # TLS solutions
    if tls_version and "1.2" in tls_version:
        print("\n✅ Your TLS version appears to be compatible (TLS 1.2 or higher).")
    else:
        print("\n⚠ TLS version may be causing issues. Try these solutions:")
        print("""
1. Add this code to your scripts to force TLS 1.2:

import requests
import urllib3
import ssl
import certifi

session = requests.Session()
session.mount('https://', urllib3.PoolManager(
    cert_reqs='CERT_REQUIRED',
    ca_certs=certifi.where(),
    ssl_version=ssl.PROTOCOL_TLSv1_2
))

# Then use session.get() instead of requests.get()
response = session.get(url, proxies={'http': None, 'https': None})
        """)
        
        print("\n2. Upgrade your OpenSSL installation:")
        print("   - For Windows, reinstall Python with the latest version")
        print("   - For Linux: sudo apt-get update && sudo apt-get install openssl")
        print("   - For macOS: brew update && brew install openssl")
    
    # Firewall solutions
    http_blocked = not port_results.get(80, False)
    https_blocked = not port_results.get(443, False)
    python_ua_blocked = not ua_results.get("python-requests/2.28.1", False)
    browser_ua_success = ua_results.get("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36", False)
    
    if http_blocked or https_blocked:
        print("\n⚠ Firewall appears to be blocking web traffic on standard ports.")
        print("Solutions:")
        print("1. Check your firewall settings and allow Python to access the internet")
        print("2. Check antivirus software that might be blocking Python connections")
        print("3. Try running Python as administrator")
    
    if python_ua_blocked and browser_ua_success:
        print("\n⚠ Your firewall/network appears to be blocking Python's default User-Agent.")
        print("Solution: Use a browser-like User-Agent in your requests:")
        print("""
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
}
response = requests.get(url, headers=headers, proxies={'http': None, 'https': None})
        """)
    
    # Final combined solution
    print("\n" + "=" * 60)
    print(" RECOMMENDED COMPLETE SOLUTION ".center(60, "="))
    print("=" * 60)
    print("""
# Use this template for all your web requests:

import requests
import urllib3
import ssl
import certifi

def get_url(url):
    # Set up a browser-like User-Agent
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
    
    # Create a session with specific TLS settings
    session = requests.Session()
    
    # Disable proxies explicitly
    proxies = {'http': None, 'https': None}
    
    try:
        # First try with default settings but our headers and no proxy
        response = session.get(url, headers=headers, proxies=proxies, timeout=10)
        return response
    except:
        try:
            # If that fails, try with explicit TLS 1.2
            adapter = urllib3.PoolManager(
                cert_reqs='CERT_REQUIRED',
                ca_certs=certifi.where(),
                ssl_version=ssl.PROTOCOL_TLSv1_2
            )
            session.mount('https://', adapter)
            response = session.get(url, headers=headers, proxies=proxies, timeout=10)
            return response
        except Exception as e:
            print(f"Error: {e}")
            return None

# Usage example
response = get_url('https://www.google.com')
if response and response.status_code == 200:
    print("Success!")
    # Process your response here
else:
    print("Failed to connect")
""")

def main():
    print("=" * 60)
    print(" TLS AND FIREWALL DIAGNOSTIC TOOL ".center(60, "="))
    print("=" * 60)
    print("Running tests to diagnose TLS and firewall issues...\n")
    
    # Run the tests
    tls_version = check_tls_compatibility()
    port_results, ua_results = check_firewall_interference()
    
    # Provide solutions
    provide_solutions(tls_version, port_results, ua_results)
    
    print("\nDiagnostic complete! Use the recommended solution in your web scraping scripts.")

if __name__ == "__main__":
    main()
```

This script will:

1. Test your TLS configuration and compatibility
2. Check if your firewall is blocking specific ports or Python requests
3. Test if using a browser-like User-Agent helps bypass restrictions
4. Provide specific solutions based on the test results
5. Give you a complete, robust template to use in all your web scraping scripts

Save this script as `tls_firewall_fix.py` and run it. It should give you a comprehensive diagnosis of any TLS or firewall issues and provide a complete solution that addresses all the potential problems mentioned in your diagnostic output.

The final solution template combines all three fixes (proxy settings, TLS version, and User-Agent) into one robust function that you can use in all your web scraping projects going forward.