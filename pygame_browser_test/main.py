import pygame
import sys

pygame.init()

# Screen size
WIDTH, HEIGHT = 400, 300
screen = pygame.display.set_mode((WIDTH, HEIGHT))
pygame.display.set_caption("Pygame Browser Test")

# Colors
BLACK = (0, 0, 0)
RED = (255, 0, 0)

# Circle position
x, y = WIDTH // 2, HEIGHT // 2
radius = 30
speed = 2

clock = pygame.time.Clock()

running = True
while running:
    for event in pygame.event.get():
        if event.type == pygame.QUIT:
            running = False

    # Move circle
    x += speed
    if x + radius > WIDTH or x - radius < 0:
        speed *= -1

    # Draw
    screen.fill(BLACK)
    pygame.draw.circle(screen, RED, (x, y), radius)
    pygame.display.flip()
    clock.tick(60)

pygame.quit()
sys.exit()
