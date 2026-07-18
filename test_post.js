const url = "https://script.google.com/macros/s/AKfycbwdCVjE-uxjDZmqsByykPds_TXQGYlq0dV6V8QyLd_YXYub9NK_u1S1QQbPgFITyiOe/exec";

async function test() {
    const payload = { action: 'addCourse', data: { Name: "C++", "Number of Sessions": 5 } };
    try {
        const res = await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify(payload)
        });
        const data = await res.json();
        console.log("Success:", data);
    } catch (e) {
        console.log("Error:", e);
    }
}
test();
