/* mobai训练 · 樱花粒子背景
 * 从模板 test.html 里抽出来的独立文件，行为保持不变。
 */
(function () {
  'use strict';

  var canvas = document.getElementById('sakuraCanvas');
  if (!canvas) return;
  var ctx = canvas.getContext('2d');
  if (!ctx) return;

  function resizeCanvas() {
    canvas.width = window.innerWidth;
    canvas.height = window.innerHeight;
  }

  resizeCanvas();
  window.addEventListener('resize', resizeCanvas);

  function Petal() {
    this.reset();
  }

  Petal.prototype.reset = function () {
    this.x = Math.random() * canvas.width;
    this.y = -10;
    this.size = Math.random() * 6 + 3;
    this.speedY = Math.random() * 0.5 + 0.3;
    this.speedX = Math.random() * 0.3 - 0.15;
    this.rotate = Math.random() * Math.PI * 2;
    this.rotateSpeed = (Math.random() - 0.5) * 0.025;
  };

  Petal.prototype.update = function () {
    this.y += this.speedY;
    this.x += this.speedX;
    this.rotate += this.rotateSpeed;
    if (this.y > canvas.height + 20) this.reset();
  };

  Petal.prototype.draw = function () {
    ctx.save();
    ctx.translate(this.x, this.y);
    ctx.rotate(this.rotate);
    ctx.fillStyle = 'rgba(255, 180, 200, 0.55)';
    ctx.beginPath();
    ctx.ellipse(0, 0, this.size, this.size * 0.55, 0, 0, Math.PI * 2);
    ctx.fill();
    ctx.restore();
  };

  var petals = [];
  var petalCount = Math.min(45, Math.floor(window.innerWidth / 22));
  for (var i = 0; i < petalCount; i++) petals.push(new Petal());

  function animate() {
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    for (var i = 0; i < petals.length; i++) {
      petals[i].update();
      petals[i].draw();
    }
    requestAnimationFrame(animate);
  }

  animate();
})();